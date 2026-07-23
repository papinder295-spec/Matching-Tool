"""
=====================================================================================
LEGAL ENTITY MATCHER - DESKTOP TOOL (Country-agnostic, GUI based)
=====================================================================================
Tkinter based desktop application for matching Direct IDs with Good IDs.

IMPORTANT FIXES INCLUDED IN THIS VERSION
-------------------------------------------------------------------------------------
1. Selected entity/legal-form patterns are REMOVED from Cleaned Name.
2. Suffix normalization is supported, and legal-form removal runs again AFTER suffix normalization.
3. Number/Roman strict validation.
4. Core-name strict validation.
5. Audit columns are exported.
6. UI multi-selection fix.
7. [NEW] Integrated 'country_city_regex_fuzzy_mapping 6.xlsx' for dynamic geo regex cleaning.
8. [NEW] Multi-Language & Spelling Standardization pipeline for both Names and Addresses.

-------------------------------------------------------------------------------------
DATA FILES - keep in same folder as this script
-------------------------------------------------------------------------------------
  - legal_form_mapping.xlsx        -> Sheet "in"
  - Universal_Suffix_Map.xlsx.xlsx -> Sheet "Sheet3", "Suffix_Map", "Mandatory_Words"
  - country_city_regex_fuzzy_mapping 6.xlsx -> Dynamic Regex removal (NEW)

-------------------------------------------------------------------------------------
SETUP
-------------------------------------------------------------------------------------
  pip install pandas sqlalchemy pyodbc thefuzz python-Levenshtein openpyxl rapidfuzz
=====================================================================================
"""

import os
import re
import sys
import threading
import traceback
from functools import lru_cache

import pandas as pd

try:
    from sqlalchemy import create_engine
except ImportError:
    create_engine = None

try:
    from thefuzz import fuzz
except ImportError:
    fuzz = None

try:
    from rapidfuzz import process as rf_process, fuzz as rf_fuzz
except ImportError:
    rf_process = None
    rf_fuzz = None

import tkinter as tk
from tkinter import ttk, messagebox, filedialog


# =====================================================================================
# 1. CONFIGURATION -- DATABASE + SCHEMA + FILES
# =====================================================================================

DB_CONFIG = {
    "SQL_SERVER": "QIG-WXRELADB501.analytics.moodys.net",
    "DATABASE": "bvdaffils",
    "ODBC_DRIVER": "ODBC Driver 17 for SQL Server",
}

SCHEMA_CONFIG = {
    "ID_COL": "Id",
    "NAME_COL": "Name",
    "ADDRESS_COL": "Address",
    "CITY_COL": "PCCity",
    "NRLINKS_COL": "NrLinks",
    "BRANCH_COL": "Branch",
    "FOREIGN_COL": "[Foreign]",
    "PERSON_COL": "Person",
    "SOURCENR_COL": "SourceNr",
    "TABLE_NAME": "Companies",
    "LINKS_TABLE": "links",
}

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

LEGAL_FORM_SOURCES = [
    (os.path.join(SCRIPT_DIR, "legal_form_mapping.xlsx"), "in"),
    (os.path.join(SCRIPT_DIR, "Universal_Suffix_Map.xlsx.xlsx"), "Sheet3"),
]

SUFFIX_MAP_SOURCE = (
    os.path.join(SCRIPT_DIR, "Universal_Suffix_Map.xlsx.xlsx"),
    "Suffix_Map",
)

MANDATORY_WORDS_SOURCE = (
    os.path.join(SCRIPT_DIR, "Universal_Suffix_Map.xlsx.xlsx"),
    "Mandatory_Words",
)

# [NEW] File 6 Path
COUNTRY_CITY_SOURCE = os.path.join(
    SCRIPT_DIR, "country_city_regex_fuzzy_mapping 6.xlsx"
)

# [NEW] Multi-Language Dictionary
STATIC_GEO_STANDARDIZATION_MAP = {
    # Italy
    r"\b(roma|rome|rom)\b": "rome",
    r"\b(milano|milan)\b": "milan",
    r"\b(torino|turin)\b": "turin",
    r"\b(venezia|venice)\b": "venice",
    r"\b(firenze|florence)\b": "florence",
    r"\b(italia|italy|italien)\b": "italy",
    # Germany
    r"\b(deutschland|germany|allemagne|germania)\b": "germany",
    r"\b(münchen|munchen|munich)\b": "munich",
    r"\b(köln|koln|cologne)\b": "cologne",
    # Spain
    r"\b(españa|espana|spain|spanien)\b": "spain",
    r"\b(sevilla|seville)\b": "seville",
    # France
    r"\b(paris|pari)\b": "paris",
    r"\b(france|frankreich|francia)\b": "france",
    # US / UK Variants
    r"\b(united states|usa|u\.s\.a\.|america|us)\b": "usa",
    r"\b(united kingdom|uk|u\.k\.|britain|great britain)\b": "uk",
}

DEFAULT_THRESHOLD = 88
DEFAULT_ADDRESS_THRESHOLD = 70


# =====================================================================================
# 2. DATABASE LOADERS (LEGAL FORMS, SUFFIXES, GEO)
# =====================================================================================

def _load_single_legal_form_sheet(file_path, sheet_name):
    if not os.path.exists(file_path):
        return pd.DataFrame(columns=["LegalForm", "Regex"])

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        return pd.DataFrame(columns=["LegalForm", "Regex"])

    if df.empty:
        return pd.DataFrame(columns=["LegalForm", "Regex"])

    df.columns = [str(c).strip() for c in df.columns]

    name_col = next(
        (c for c in df.columns if "legal" in c.lower() and "form" in c.lower()),
        next((c for c in df.columns if c.lower() == "legalform"), df.columns[0]),
    )
    regex_col = next((c for c in df.columns if "regex" in c.lower()), None)

    out = pd.DataFrame()
    out["LegalForm"] = df[name_col].astype(str).str.strip()
    out["Regex"] = df[regex_col].astype(str).str.strip() if regex_col else ""
    out = out[
        out["LegalForm"].notna()
        & (out["LegalForm"] != "")
        & (out["LegalForm"].str.lower() != "nan")
    ]
    return out


def load_legal_form_database(sources=None):
    if sources is None:
        sources = LEGAL_FORM_SOURCES

    frames = []
    for file_path, sheet_name in sources:
        df = _load_single_legal_form_sheet(file_path, sheet_name)
        if not df.empty:
            frames.append(df)

    if not frames:
        return pd.DataFrame(columns=["LegalForm", "Regex"])

    merged = pd.concat(frames, ignore_index=True)
    merged = merged.drop_duplicates(subset=["LegalForm"], keep="first")
    merged = merged.sort_values("LegalForm").reset_index(drop=True)
    return merged


def load_suffix_map(source=None):
    if source is None:
        source = SUFFIX_MAP_SOURCE

    file_path, sheet_name = source
    if not os.path.exists(file_path):
        return {}

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        return {}

    if df.empty:
        return {}

    df.columns = [str(c).strip() for c in df.columns]

    suffix_col = next((c for c in df.columns if "suffix" in c.lower()), df.columns[0])
    replacement_col = next((c for c in df.columns if "replace" in c.lower()), None)

    if replacement_col is None:
        return {}

    suffix_map = {}
    for _, row in df.iterrows():
        suffix = str(row[suffix_col]).strip().lower()
        replacement = str(row[replacement_col]).strip().lower()

        if suffix in ("", "nan", "common/global") or replacement in ("", "nan"):
            continue
        suffix_map[suffix] = replacement

    return suffix_map


def load_mandatory_words(source=None):
    if source is None:
        source = MANDATORY_WORDS_SOURCE

    file_path, sheet_name = source
    if not os.path.exists(file_path):
        return set()

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception:
        return set()

    if df.empty:
        return set()

    df.columns = [str(c).strip() for c in df.columns]
    word_col = df.columns[0]

    words = set()
    for val in df[word_col].dropna():
        w = str(val).strip().lower()
        if w and w != "nan":
            words.add(w)

    return words


# [NEW] Load Geo Standardization Regex
def build_geo_standardizer(mapping_dict):
    compiled_rules = []
    for pattern, target in mapping_dict.items():
        compiled_rules.append((re.compile(pattern, flags=re.IGNORECASE), target))
    return compiled_rules


# [NEW] Load Country/City file patterns
def load_country_city_regex(file_path=COUNTRY_CITY_SOURCE):
    if not os.path.exists(file_path):
        return None

    try:
        df = pd.read_excel(file_path)
        # Select 'RegexFuzzy' column specifically to avoid matching 'RegexPart'
        regex_cols = [c for c in df.columns if "regex" in c.lower() and "fuzzy" in c.lower()]

        if not regex_cols:
            # Fallback to the first column containing regex if fuzzy not found
            regex_cols = [c for c in df.columns if "regex" in c.lower()]

        if not regex_cols:
            return None

        patterns = df[regex_cols[0]].dropna().astype(str).str.strip().tolist()
        valid_patterns = [p for p in patterns if p and p.lower() != "nan" and p != "-"]

        if valid_patterns:
            combined_pattern = "|".join([f"(?:{p})" for p in valid_patterns])
            return re.compile(combined_pattern, flags=re.IGNORECASE)
    except Exception:
        pass
    return None


# =====================================================================================
# 3. REGEX + CLEANING HELPERS
# =====================================================================================

def normalize_text_for_regex(text):
    text = str(text).lower().strip()
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def get_suffix_related_terms(term, suffix_map):
    terms = set()
    term_norm = normalize_text_for_regex(term)
    if term_norm:
        terms.add(term_norm)

    if not suffix_map or not term_norm:
        return terms

    replacement = suffix_map.get(term_norm)
    if replacement:
        rep_norm = normalize_text_for_regex(replacement)
        if rep_norm:
            terms.add(rep_norm)

    for k, v in suffix_map.items():
        k_norm = normalize_text_for_regex(k)
        v_norm = normalize_text_for_regex(v)
        if term_norm == v_norm and k_norm:
            terms.add(k_norm)
        if term_norm == k_norm and v_norm:
            terms.add(v_norm)

    return terms


def word_phrase_regex(phrase):
    phrase = normalize_text_for_regex(phrase)
    if not phrase:
        return None
    parts = [re.escape(p) for p in phrase.split() if p]
    if not parts:
        return None
    return r"\b" + r"\s+".join(parts) + r"\b"


def patterns_to_regex(selected_forms, custom_words, legal_form_df, suffix_map=None):
    parts = []
    added_simple_terms = set()

    def add_simple_term(term):
        for t in get_suffix_related_terms(term, suffix_map):
            if not t or t in added_simple_terms:
                continue
            pattern = word_phrase_regex(t)
            if pattern:
                parts.append(pattern)
                added_simple_terms.add(t)

    for form in selected_forms:
        form_str = str(form).strip()
        if not form_str:
            continue
        add_simple_term(form_str)

        row = legal_form_df[legal_form_df["LegalForm"].str.upper() == form_str.upper()]
        if not row.empty and str(row.iloc[0]["Regex"]).strip() not in ("", "nan"):
            pattern = str(row.iloc[0]["Regex"]).strip()
            try:
                re.compile(pattern)
                parts.append(f"(?:{pattern})")
            except re.error:
                pass

    for word in custom_words:
        word = str(word).strip()
        if not word:
            continue
        add_simple_term(word)

    if not parts:
        return None

    parts = sorted(parts, key=len, reverse=True)
    return re.compile("|".join(parts), flags=re.IGNORECASE)


# [UPDATED] Clean Name pipeline with Geo Standardization
def make_clean_name_fn(entity_regex, suffix_map=None, mandatory_words=None, geo_rules=None, dynamic_geo_regex=None):
    if suffix_map is None:
        suffix_map = {}
    if mandatory_words is None:
        mandatory_words = set()
    if geo_rules is None:
        geo_rules = []

    if suffix_map:
        sorted_suffixes = sorted(suffix_map.keys(), key=len, reverse=True)
        suffix_patterns = []
        for s in sorted_suffixes:
            s_norm = normalize_text_for_regex(s)
            if not s_norm:
                continue
            pattern = word_phrase_regex(s_norm)
            if pattern:
                suffix_patterns.append(pattern)
        suffix_regex = re.compile("|".join(suffix_patterns), flags=re.IGNORECASE) if suffix_patterns else None
    else:
        suffix_regex = None

    normalized_suffix_map = {
        normalize_text_for_regex(k): normalize_text_for_regex(v)
        for k, v in suffix_map.items()
        if normalize_text_for_regex(k) and normalize_text_for_regex(v)
    }

    def remove_entity_patterns(text):
        if entity_regex is not None and text:
            text = entity_regex.sub(" ", text)
            text = re.sub(r"\s+", " ", text).strip()
        return text

    @lru_cache(maxsize=None)
    def clean_name(name):
        if pd.isna(name) or not str(name).strip():
            return ""

        original = str(name).lower().strip()
        cleaned = original
        cleaned = re.sub(r"\s+", " ", cleaned).strip()

        # [NEW] Geo Rule Standardization
        for rule_regex, standard_term in geo_rules:
            cleaned = rule_regex.sub(standard_term, cleaned)
            
        # [NEW] Dynamic Geo Regex Removal
        if dynamic_geo_regex:
            cleaned = dynamic_geo_regex.sub(" ", cleaned)

        preserved_words = set()
        if mandatory_words:
            name_words = set(re.findall(r"[a-z]+", cleaned))
            preserved_words = name_words & mandatory_words

        cleaned = remove_entity_patterns(cleaned)
        cleaned = re.sub(r"[^a-z0-9\s]", " ", cleaned)
        cleaned = re.sub(r"\s+", " ", cleaned).strip()

        if suffix_regex:
            def replace_suffix(match):
                matched_text = normalize_text_for_regex(match.group(0))
                return normalized_suffix_map.get(matched_text, matched_text)
            cleaned = suffix_regex.sub(replace_suffix, cleaned)
            cleaned = re.sub(r"\s+", " ", cleaned).strip()

        cleaned = remove_entity_patterns(cleaned)
        cleaned = re.sub(r"[^a-z0-9\s]", " ", cleaned)
        cleaned = re.sub(r"\s+", " ", cleaned).strip()

        if preserved_words:
            current_words = set(cleaned.split())
            missing_words = preserved_words - current_words
            if missing_words:
                cleaned = cleaned + " " + " ".join(sorted(missing_words))
                cleaned = re.sub(r"\s+", " ", cleaned).strip()

        return cleaned

    return clean_name


# [UPDATED] Clean Address pipeline with Geo Standardization
def make_clean_address_fn(geo_rules=None, dynamic_geo_regex=None):
    if geo_rules is None:
        geo_rules = []

    @lru_cache(maxsize=None)
    def clean_address(addr):
        if pd.isna(addr) or not addr or str(addr).strip() == "":
            return ""
        
        cleaned = str(addr).lower()
        
        # [NEW] Geo Rule Standardization (Roma -> rome etc.)
        for rule_regex, standard_term in geo_rules:
            cleaned = rule_regex.sub(standard_term, cleaned)
            
        # [NEW] Dynamic Geo Regex Removal for Addresses
        if dynamic_geo_regex:
            cleaned = dynamic_geo_regex.sub(" ", cleaned)

        cleaned = re.sub(r"[^a-z0-9\s]", " ", cleaned)
        cleaned = re.sub(r"\s+", " ", cleaned).strip()
        return cleaned
        
    return clean_address


# =====================================================================================
# 3A. NUMBER / ROMAN COUNTING STRICT VALIDATION
# =====================================================================================

ROMAN_NUMERAL_MAP = {
    "i": 1, "ii": 2, "iii": 3, "iv": 4, "v": 5,
    "vi": 6, "vii": 7, "viii": 8, "ix": 9, "x": 10,
    "xi": 11, "xii": 12, "xiii": 13, "xiv": 14, "xv": 15,
    "xvi": 16, "xvii": 17, "xviii": 18, "xix": 19, "xx": 20,
    "xxi": 21, "xxii": 22, "xxiii": 23, "xxiv": 24, "xxv": 25,
}

def _normalize_number_token(token):
    if token is None:
        return None
    token = str(token).strip().lower()
    if token.isdigit():
        try:
            return str(int(token))
        except Exception:
            return token
    if token in ROMAN_NUMERAL_MAP:
        return str(ROMAN_NUMERAL_MAP[token])
    return None

def extract_number_roman_signature(cleaned_name):
    if pd.isna(cleaned_name) or not str(cleaned_name).strip():
        return tuple()
    tokens = str(cleaned_name).lower().strip().split()
    signature = []
    for token in tokens:
        normalized = _normalize_number_token(token)
        if normalized is not None:
            signature.append(normalized)
    return tuple(signature)

def remove_number_roman_tokens(cleaned_name):
    if pd.isna(cleaned_name) or not str(cleaned_name).strip():
        return ""
    tokens = str(cleaned_name).lower().strip().split()
    base_tokens = []
    for token in tokens:
        if _normalize_number_token(token) is None:
            base_tokens.append(token)
    return "".join(base_tokens)

def number_roman_rule_status(direct_clean_name, good_clean_name):
    direct_sig = extract_number_roman_signature(direct_clean_name)
    good_sig = extract_number_roman_signature(good_clean_name)
    direct_base = remove_number_roman_tokens(direct_clean_name)
    good_base = remove_number_roman_tokens(good_clean_name)

    if direct_sig or good_sig:
        if direct_sig != good_sig:
            return False, f"Rejected: Number/Roman mismatch Direct={direct_sig} Good={good_sig}"
        if direct_base != good_base:
            return False, f"Rejected: Base name mismatch Direct='{direct_base}' Good='{good_base}'"
        return True, f"Passed: Same base name and same Number/Roman {direct_sig}"
    return True, "Passed: No Number/Roman on both sides"


# =====================================================================================
# 3B. CORE NAME STRICT VALIDATION
# =====================================================================================

BASE_CORE_NAME_IGNORE_WORDS = {
    "oy", "ab", "as", "asa", "aps", "a", "s", "ltd", "limited", "llc", "inc", "incorporated",
    "corp", "corporation", "co", "company", "companies", "plc", "sa", "spa", "srl", "gmbh", 
    "bv", "nv", "pt", "tbk", "cv", "kg", "ag", "se", "pte", "sdn", "bhd", "berhad", 
    "private", "pvt", "llp", "limite", "ltda", "sarl", "sas", "bvba", "osakeyhtio", 
    "osakeyhti", "aktiebolag", "aktiebolaget", "osuuskunta", "kommandiittiyhtio", 
    "avoin", "yhtio", "the", "and", "of", "for", "in", "on", "at", "by", "to",
}

def build_runtime_core_ignore_words(suffix_map=None):
    ignore_words = set(BASE_CORE_NAME_IGNORE_WORDS)
    if suffix_map:
        for key, value in suffix_map.items():
            for item in (key, value):
                item = normalize_text_for_regex(item)
                if not item or item == "nan":
                    continue
                for token in item.split():
                    if token:
                        ignore_words.add(token)
    return ignore_words

def extract_core_name_tokens(cleaned_name, ignore_words=None):
    if ignore_words is None:
        ignore_words = BASE_CORE_NAME_IGNORE_WORDS
    if pd.isna(cleaned_name) or not str(cleaned_name).strip():
        return set()
    tokens = str(cleaned_name).lower().strip().split()
    core_tokens = set()
    for token in tokens:
        token = token.strip().lower()
        if not token or _normalize_number_token(token) is not None or token in ignore_words or len(token) <= 1:
            continue
        core_tokens.add(token)
    return core_tokens

def core_name_rule_status(direct_clean_name, good_clean_name, ignore_words=None):
    direct_tokens = extract_core_name_tokens(direct_clean_name, ignore_words=ignore_words)
    good_tokens = extract_core_name_tokens(good_clean_name, ignore_words=ignore_words)

    if direct_tokens and good_tokens:
        common_tokens = direct_tokens & good_tokens
        if not common_tokens:
            return False, f"Rejected: Core name mismatch Direct={sorted(direct_tokens)} Good={sorted(good_tokens)}"
        return True, f"Passed: Common core token(s)={sorted(common_tokens)}"
    return True, "Passed: Core token rule skipped due to missing core tokens"


# =====================================================================================
# 4. SQL QUERY BUILDERS
# =====================================================================================

def build_company_query(good_pattern, direct_pattern, schema):
    id_col = schema["ID_COL"]
    name_col = schema["NAME_COL"]
    addr_col = schema["ADDRESS_COL"]
    city_col = schema["CITY_COL"]
    nrlinks_col = schema["NRLINKS_COL"]
    branch_col = schema["BRANCH_COL"]
    foreign_col = schema["FOREIGN_COL"]
    person_col = schema["PERSON_COL"]
    sourcenr_col = schema["SOURCENR_COL"]
    table = schema["TABLE_NAME"]

    select_cols = (
        f"{id_col}, {name_col}, {addr_col}, {city_col}, "
        f"{nrlinks_col}, {branch_col}, {foreign_col}, {sourcenr_col}"
    )

    fixed_filters = f"""
      AND ({person_col} NOT IN ('I', 'L', 'K', 'D') OR {person_col} IS NULL)
      AND ({branch_col} IS NULL OR {branch_col} IN (0, 3))
      AND ({foreign_col} IS NULL OR {foreign_col} IN (0, 3))
    """

    direct_base = direct_pattern.rstrip("%")

    query = f"""
WITH GoodCompanies AS (
    SELECT {select_cols}
    FROM {table}
    WHERE
    ({id_col} LIKE '{good_pattern}%' AND {id_col} NOT LIKE '{direct_base}%')
    {fixed_filters}
),
DirectCompanies AS (
    SELECT {select_cols}
    FROM {table}
    WHERE
    ({id_col} LIKE '{direct_pattern}%')
    {fixed_filters}
)
SELECT 'Good' AS Source, {select_cols} FROM GoodCompanies
UNION ALL
SELECT 'Direct' AS Source, {select_cols} FROM DirectCompanies;
"""
    return query


def build_links_query(good_pattern, direct_pattern, schema):
    id_like_a = good_pattern
    id_like_b = direct_pattern.rstrip("%")
    links_table = schema["LINKS_TABLE"]
    return (
        f"SELECT ParentId, ChildId, UO FROM {links_table} "
        f"WHERE ParentId LIKE '{id_like_a}%' OR ChildId LIKE '{id_like_a}%' "
        f"OR ParentId LIKE '{id_like_b}%' OR ChildId LIKE '{id_like_b}%'"
    )


# =====================================================================================
# 5. MATCHING PIPELINE
# =====================================================================================

def run_matching_pipeline(params, log_fn):
    if create_engine is None or fuzz is None:
        raise RuntimeError("Required libraries missing. 'pip install pandas sqlalchemy pyodbc thefuzz python-Levenshtein openpyxl rapidfuzz'")

    schema = SCHEMA_CONFIG
    conn_str = (
        f"mssql+pyodbc://@{DB_CONFIG['SQL_SERVER']}/{DB_CONFIG['DATABASE']}"
        f"?driver={DB_CONFIG['ODBC_DRIVER'].replace(' ', '+')}&trusted_connection=yes"
    )

    log_fn("Connecting to database...")
    engine = create_engine(conn_str)

    company_query = build_company_query(params["good_pattern"], params["direct_pattern"], schema)
    links_query = build_links_query(params["good_pattern"], params["direct_pattern"], schema)

    log_fn("Fetching company records...")
    df = pd.read_sql(company_query, engine)

    use_hierarchy = params.get("use_hierarchy", True)
    parent_map, child_map = {}, {}
    if use_hierarchy:
        log_fn("Fetching hierarchy (links) data...")
        try:
            links_df = pd.read_sql(links_query, engine)
            parent_map = links_df.groupby("ParentId")["ChildId"].apply(set).to_dict()
            child_map = links_df.groupby("ChildId")["ParentId"].apply(set).to_dict()
        except Exception as e:
            log_fn(f"[WARNING] Links table fetch fail hui: {e}")

    id_col = schema["ID_COL"]
    name_col = schema["NAME_COL"]
    addr_col = schema["ADDRESS_COL"]
    city_col = schema["CITY_COL"]
    nrlinks_col = schema["NRLINKS_COL"]
    sourcenr_col = schema["SOURCENR_COL"]

    clean_name = make_clean_name_fn(
        params["entity_regex"],
        suffix_map=params.get("suffix_map"),
        mandatory_words=params.get("mandatory_words"),
        geo_rules=params.get("geo_rules"),
        dynamic_geo_regex=params.get("dynamic_geo_regex")
    )
    
    clean_addr_fn = make_clean_address_fn(
        geo_rules=params.get("geo_rules"),
        dynamic_geo_regex=params.get("dynamic_geo_regex")
    )
    
    core_ignore_words = build_runtime_core_ignore_words(params.get("suffix_map"))

    log_fn("Cleaning names & addresses...")
    df["Cleaned Name"] = df[name_col].map(clean_name)
    df["Cleaned Address"] = (
        df[addr_col].fillna("").astype(str) + " " + df[city_col].fillna("").astype(str)
    ).map(clean_addr_fn)
    df["Has_Address"] = df["Cleaned Address"].str.strip() != ""

    df["Number_Roman_Signature"] = df["Cleaned Name"].map(extract_number_roman_signature)
    df["Base_Name_Without_Number_Roman"] = df["Cleaned Name"].map(remove_number_roman_tokens)
    df["Core_Name_Tokens"] = df["Cleaned Name"].map(
        lambda x: sorted(extract_core_name_tokens(x, ignore_words=core_ignore_words))
    )

    good_df = df[df["Source"] == "Good"].copy()
    direct_df = df[df["Source"] == "Direct"].copy()

    good_df["First_Word"] = good_df["Cleaned Name"].apply(lambda x: x.split()[0] if x else None)
    direct_df["First_Word"] = direct_df["Cleaned Name"].apply(lambda x: x.split()[0] if x else None)
    good_groups = good_df.groupby("First_Word")

    threshold = params["threshold"]
    addr_threshold = params["address_threshold"]

    all_matches = []
    no_match_list = []
    total = len(direct_df)

    log_fn(f"Matching {total} Direct records against {len(good_df)} Good records...")

    for i, row in enumerate(direct_df.itertuples(index=False), start=1):
        row_d = dict(zip(direct_df.columns, row))
        d_id = row_d[id_col]
        d_name = row_d[name_col]
        c_name = row_d["Cleaned Name"]
        c_addr = row_d["Cleaned Address"]
        has_addr = row_d["Has_Address"]
        d_sn = row_d[sourcenr_col]
        d_nl = row_d[nrlinks_col]
        d_fw = row_d["First_Word"]

        if i % 200 == 0:
            log_fn(f"...{i}/{total} processed")

        if not d_fw or d_fw not in good_groups.groups:
            no_match_list.append({
                "Direct ID": d_id, "Direct Original Name": d_name, "Direct SourceNr": d_sn, "Direct Cleaned Name": c_name,
                "No Match Reason": "No candidate group found by first clean word",
            })
            continue

        candidates = good_groups.get_group(d_fw).copy()
        candidates = candidates[candidates[id_col] != d_id]

        if candidates.empty:
            no_match_list.append({
                "Direct ID": d_id, "Direct Original Name": d_name, "Direct SourceNr": d_sn, "Direct Cleaned Name": c_name,
                "No Match Reason": "No candidate after excluding same ID",
            })
            continue

        number_rule_results = candidates["Cleaned Name"].apply(lambda x: number_roman_rule_status(c_name, x))
        candidates["Number_Roman_Rule_Passed"] = number_rule_results.apply(lambda x: x[0])
        candidates["Number_Roman_Rule_Message"] = number_rule_results.apply(lambda x: x[1])
        candidates = candidates[candidates["Number_Roman_Rule_Passed"] == True]

        if candidates.empty:
            no_match_list.append({
                "Direct ID": d_id, "Direct Original Name": d_name, "Direct SourceNr": d_sn, "Direct Cleaned Name": c_name,
                "No Match Reason": "All candidates rejected due to Number/Roman/Base-name rule",
            })
            continue

        core_rule_results = candidates["Cleaned Name"].apply(lambda x: core_name_rule_status(c_name, x, ignore_words=core_ignore_words))
        candidates["Core_Name_Rule_Passed"] = core_rule_results.apply(lambda x: x[0])
        candidates["Core_Name_Rule_Message"] = core_rule_results.apply(lambda x: x[1])
        candidates = candidates[candidates["Core_Name_Rule_Passed"] == True]

        if candidates.empty:
            no_match_list.append({
                "Direct ID": d_id, "Direct Original Name": d_name, "Direct SourceNr": d_sn, "Direct Cleaned Name": c_name,
                "No Match Reason": "All candidates rejected due to Core-name mismatch rule",
            })
            continue

        cand_names = candidates["Cleaned Name"].tolist()
        if rf_process is not None:
            candidates["Name_Score"] = rf_process.cdist([c_name], cand_names, scorer=rf_fuzz.ratio)[0]
        else:
            candidates["Name_Score"] = [fuzz.ratio(c_name, x) for x in cand_names]

        if has_addr:
            cand_addrs = candidates["Cleaned Address"].tolist()
            cand_has_addr = candidates["Has_Address"].tolist()
            if rf_process is not None:
                raw_addr_scores = rf_process.cdist([c_addr], cand_addrs, scorer=rf_fuzz.ratio)[0]
            else:
                raw_addr_scores = [fuzz.ratio(c_addr, x) for x in cand_addrs]
            candidates["Address_Score"] = [score if has else None for score, has in zip(raw_addr_scores, cand_has_addr)]
        else:
            candidates["Address_Score"] = None

        def passes(r):
            if r["Name_Score"] < threshold: return False
            if has_addr and r["Has_Address"]: return r["Address_Score"] is not None and r["Address_Score"] >= addr_threshold
            return True

        found = candidates[candidates.apply(passes, axis=1)]

        if found.empty:
            no_match_list.append({
                "Direct ID": d_id, "Direct Original Name": d_name, "Direct SourceNr": d_sn, "Direct Cleaned Name": c_name,
                "No Match Reason": "No candidate passed fuzzy/address threshold after strict rules",
            })
            continue

        for _, m in found.iterrows():
            all_matches.append({
                "Direct ID": d_id, "Direct Original Name": d_name, "Direct SourceNr": d_sn,
                "Matched Good ID": m[id_col], "Matched Good Original Name": m[name_col], "Matched SourceNr": m[sourcenr_col],
                "Name Fuzzy Score": m["Name_Score"], "Address Fuzzy Score": m["Address_Score"] if pd.notna(m["Address_Score"]) else "N/A (no address)",
                "Direct NrLinks": d_nl, "Matched NrLinks": m[nrlinks_col],
                "Direct Cleaned Name": c_name, "Matched Cleaned Name": m["Cleaned Name"],
                "Direct Number/Roman Signature": extract_number_roman_signature(c_name), "Matched Number/Roman Signature": extract_number_roman_signature(m["Cleaned Name"]),
                "Number/Roman Validation": m.get("Number_Roman_Rule_Message", ""),
                "Direct Core Tokens": sorted(extract_core_name_tokens(c_name, ignore_words=core_ignore_words)),
                "Matched Core Tokens": sorted(extract_core_name_tokens(m["Cleaned Name"], ignore_words=core_ignore_words)),
                "Core Name Validation": m.get("Core_Name_Rule_Message", ""),
            })

    log_fn("Consolidating matches (hierarchy validation + best match selection)...")
    all_matches_df = pd.DataFrame(all_matches)
    single_match_list = []
    multi_match_list = []

    if not all_matches_df.empty:
        for did, group in all_matches_df.groupby("Direct ID"):
            filtered = apply_hierarchy_filter(group, parent_map, child_map)
            if len(filtered) > 1:
                filtered = apply_address_filter(filtered, addr_threshold)
            if len(filtered) == 1:
                single_match_list.append(filtered.iloc[0].to_dict())
            else:
                multi_match_list.append(filtered)

    single_match_df = pd.DataFrame(single_match_list)
    multi_match_df = pd.concat(multi_match_list, ignore_index=True) if multi_match_list else pd.DataFrame()

    final_export_cols = [
        "Direct ID", "Direct Original Name", "Direct SourceNr", "Matched Good ID", "Matched Good Original Name", "Matched SourceNr",
        "Name Fuzzy Score", "Address Fuzzy Score", "Validation_Message", "Direct NrLinks", "Matched NrLinks",
        "Direct Cleaned Name", "Matched Cleaned Name", "Direct Number/Roman Signature", "Matched Number/Roman Signature",
        "Number/Roman Validation", "Direct Core Tokens", "Matched Core Tokens", "Core Name Validation",
    ]

    def get_clean_df(df_in):
        return df_in.reindex(columns=final_export_cols)

    if not single_match_df.empty:
        hierarchy_matches = single_match_df[single_match_df["Validation_Message"].str.contains("Link", na=False)]
        address_matches = single_match_df[single_match_df["Validation_Message"].str.contains("Address/City", na=False)]
        fuzzy_matches = single_match_df[~single_match_df["Validation_Message"].str.contains("Link", na=False) & ~single_match_df["Validation_Message"].str.contains("Address/City", na=False)]

        hierarchy_out = get_clean_df(hierarchy_matches)
        address_out = get_clean_df(address_matches)
        fuzzy_out = get_clean_df(fuzzy_matches)
        high_score_no_link = get_clean_df(single_match_df[(single_match_df["Name Fuzzy Score"] >= 95) & (single_match_df["Direct NrLinks"] == 0)])
    else:
        hierarchy_out = address_out = fuzzy_out = high_score_no_link = pd.DataFrame(columns=final_export_cols)

    multi_match_out = get_clean_df(multi_match_df) if not multi_match_df.empty else pd.DataFrame(columns=final_export_cols)

    no_match_cols = ["Direct ID", "Direct Original Name", "Direct SourceNr", "Direct Cleaned Name", "No Match Reason"]
    no_match_df = pd.DataFrame(no_match_list).reindex(columns=no_match_cols) if no_match_list else pd.DataFrame(columns=no_match_cols)

    output_path = params["output_path"]
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    log_fn(f"Saving output to: {output_path}")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        hierarchy_out.to_excel(writer, sheet_name="Hierarchy Matches", index=False)
        address_out.to_excel(writer, sheet_name="Address-City Matches", index=False)
        fuzzy_out.to_excel(writer, sheet_name="Fuzzy Matches", index=False)
        high_score_no_link.to_excel(writer, sheet_name="High Score No Link", index=False)
        multi_match_out.to_excel(writer, sheet_name="Multi Match", index=False)
        no_match_df.to_excel(writer, sheet_name="No Match", index=False)

    log_fn("Done!")
    return output_path


# =====================================================================================
# 5A. CONSOLIDATION FILTERS
# =====================================================================================

def apply_hierarchy_filter(group, parent_map, child_map):
    did = group["Direct ID"].iloc[0]
    if did in parent_map:
        matches = group[group["Matched Good ID"].apply(lambda x: x in parent_map[did] or bool(parent_map.get(x, set()) & parent_map[did]))]
        if not matches.empty: return matches.assign(Validation_Message="Identified by Subsidiary Link")
    if did in child_map:
        matches = group[group["Matched Good ID"].apply(lambda x: x in child_map[did] or bool(child_map.get(x, set()) & child_map[did]))]
        if not matches.empty: return matches.assign(Validation_Message="Identified by Shareholder Link")
    if len(group) == 1: return group.assign(Validation_Message="Single Match (No Hierarchy)")
    return group.assign(Validation_Message="Multiple Matches (No Hierarchy)")

def apply_address_filter(group, addr_threshold):
    if len(group) <= 1: return group
    addr_scores_numeric = pd.to_numeric(group["Address Fuzzy Score"], errors="coerce")
    qualifying = group[addr_scores_numeric >= addr_threshold]
    if len(qualifying) == 1: return qualifying.assign(Validation_Message="Identified by Address/City Match")
    return group


# =====================================================================================
# 6. GUI
# =====================================================================================

class MatcherApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Legal Entity Matcher - Good ID / Direct ID Matching Tool")
        self.geometry("780x680")
        self.resizable(False, False)

        self.legal_form_df = load_legal_form_database()
        self.suffix_map = load_suffix_map()
        self.mandatory_words = load_mandatory_words()
        
        # [NEW] Load Geo configurations
        self.compiled_geo_rules = build_geo_standardizer(STATIC_GEO_STANDARDIZATION_MAP)
        self.dynamic_geo_regex = load_country_city_regex()

        self.output_dir = os.path.join(os.path.expanduser("~"), "Desktop")

        self._build_ui()

        self._log(f"Legal Form Database loaded: {len(self.legal_form_df)} entries")
        self._log(f"Suffix Map loaded: {len(self.suffix_map)} normalization rules")
        self._log(f"Mandatory Words loaded: {len(self.mandatory_words)} words")
        self._log(f"Static Geo Standardization rules loaded: {len(self.compiled_geo_rules)} rules")
        if self.dynamic_geo_regex:
            self._log(f"Dynamic Geo Regex loaded successfully from File 6.")
        else:
            self._log(f"[WARNING] Dynamic Geo Regex file (File 6) not found or empty.")

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        frm_top = ttk.Frame(self)
        frm_top.pack(fill="x", **pad)

        ttk.Label(frm_top, text="Country Name:").grid(row=0, column=0, sticky="w")
        self.country_var = tk.StringVar()
        ttk.Entry(frm_top, textvariable=self.country_var, width=25).grid(row=0, column=1, sticky="w", padx=5)

        ttk.Label(frm_top, text="Good ID Pattern:").grid(row=1, column=0, sticky="w", pady=4)
        self.good_id_var = tk.StringVar()
        ttk.Entry(frm_top, textvariable=self.good_id_var, width=25).grid(row=1, column=1, sticky="w", padx=5)
        ttk.Label(frm_top, text="e.g. FIC").grid(row=1, column=2, sticky="w")

        ttk.Label(frm_top, text="Direct ID Pattern:").grid(row=2, column=0, sticky="w", pady=4)
        self.direct_id_var = tk.StringVar()
        ttk.Entry(frm_top, textvariable=self.direct_id_var, width=25).grid(row=2, column=1, sticky="w", padx=5)
        ttk.Label(frm_top, text="e.g. FIC*").grid(row=2, column=2, sticky="w")

        ttk.Label(frm_top, text="Fuzzy Threshold (Name):").grid(row=3, column=0, sticky="w", pady=4)
        self.threshold_var = tk.IntVar(value=DEFAULT_THRESHOLD)
        ttk.Spinbox(frm_top, from_=50, to=100, textvariable=self.threshold_var, width=5).grid(row=3, column=1, sticky="w", padx=5)

        ttk.Label(frm_top, text="Fuzzy Threshold (Address):").grid(row=4, column=0, sticky="w", pady=4)
        self.addr_threshold_var = tk.IntVar(value=DEFAULT_ADDRESS_THRESHOLD)
        ttk.Spinbox(frm_top, from_=30, to=100, textvariable=self.addr_threshold_var, width=5).grid(row=4, column=1, sticky="w", padx=5)

        self.use_hierarchy_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm_top, text="Use hierarchy (links table) validation", variable=self.use_hierarchy_var).grid(row=5, column=0, columnspan=2, sticky="w", pady=4)

        frm_entity = ttk.LabelFrame(self, text="Entity Pattern (Legal Form) - Multi Select")
        frm_entity.pack(fill="x", **pad)

        frm_search = ttk.Frame(frm_entity)
        frm_search.pack(fill="x", padx=8, pady=(8, 2))
        ttk.Label(frm_search, text="Search:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._filter_entity_list)
        ttk.Entry(frm_search, textvariable=self.search_var, width=30).pack(side="left", padx=5)

        ttk.Button(frm_search, text="Select Visible", command=self._select_visible_patterns).pack(side="left", padx=(10, 3))
        ttk.Button(frm_search, text="Select All", command=self._select_all_patterns).pack(side="left", padx=3)
        ttk.Button(frm_search, text="Clear", command=self._clear_all_patterns).pack(side="left", padx=3)

        self._all_forms = sorted(self.legal_form_df["LegalForm"].tolist())
        self._selected_forms_set = set()

        self.entity_listbox = tk.Listbox(frm_entity, selectmode="extended", height=8, exportselection=False)
        for form in self._all_forms:
            self.entity_listbox.insert("end", form)
        self.entity_listbox.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        self.entity_listbox.bind("<<ListboxSelect>>", self._on_entity_list_select)

        scrollbar = ttk.Scrollbar(frm_entity, orient="vertical", command=self.entity_listbox.yview)
        scrollbar.pack(side="left", fill="y")
        self.entity_listbox.config(yscrollcommand=scrollbar.set)

        frm_custom = ttk.Frame(frm_entity)
        frm_custom.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        ttk.Label(frm_custom, text="Custom Patterns\n(comma separated)\ne.g. OY, AB, OYK, LTD").pack(anchor="w")
        self.custom_pattern_var = tk.StringVar()
        ttk.Entry(frm_custom, textvariable=self.custom_pattern_var, width=35).pack(fill="x", pady=6)

        ttk.Label(frm_custom, text="Selected Patterns:").pack(anchor="w", pady=(10, 0))
        self.selected_patterns_var = tk.StringVar(value="0 selected")
        self.selected_patterns_entry = ttk.Entry(frm_custom, textvariable=self.selected_patterns_var, width=35, state="readonly")
        self.selected_patterns_entry.pack(fill="x", pady=4)

        frm_out = ttk.Frame(self)
        frm_out.pack(fill="x", **pad)
        ttk.Label(frm_out, text="Output Folder:").pack(side="left")
        self.output_dir_var = tk.StringVar(value=self.output_dir)
        ttk.Entry(frm_out, textvariable=self.output_dir_var, width=55).pack(side="left", padx=5)
        ttk.Button(frm_out, text="Browse...", command=self._browse_folder).pack(side="left")

        self.find_btn = ttk.Button(self, text="Find Matching", command=self._on_find_click)
        self.find_btn.pack(pady=10)

        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.pack(fill="x", padx=10)

        frm_log = ttk.LabelFrame(self, text="Log")
        frm_log.pack(fill="both", expand=True, **pad)
        self.log_text = tk.Text(frm_log, height=12, state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)

    def _sync_visible_selection_to_set(self):
        if not hasattr(self, "entity_listbox") or not hasattr(self, "_selected_forms_set"): return
        visible_items = [self.entity_listbox.get(i) for i in range(self.entity_listbox.size())]
        selected_indices = set(self.entity_listbox.curselection())
        for idx, form in enumerate(visible_items):
            if idx in selected_indices: self._selected_forms_set.add(form)
            else: self._selected_forms_set.discard(form)

    def _refresh_selected_patterns_display(self):
        selected = sorted(self._selected_forms_set)
        if not selected: text = "0 selected"
        elif len(selected) <= 8: text = f"{len(selected)} selected: " + ", ".join(selected)
        else: text = f"{len(selected)} selected: " + ", ".join(selected[:8]) + " ..."
        self.selected_patterns_var.set(text)

    def _reload_entity_listbox(self):
        search_term = self.search_var.get().strip().lower()
        self.entity_listbox.delete(0, "end")
        for form in self._all_forms:
            if not search_term or search_term in form.lower():
                self.entity_listbox.insert("end", form)
                if form in self._selected_forms_set: self.entity_listbox.selection_set(self.entity_listbox.size() - 1)
        self._refresh_selected_patterns_display()

    def _filter_entity_list(self, *args):
        self._sync_visible_selection_to_set()
        self._reload_entity_listbox()

    def _on_entity_list_select(self, event=None):
        self._sync_visible_selection_to_set()
        self._refresh_selected_patterns_display()

    def _select_visible_patterns(self):
        for i in range(self.entity_listbox.size()):
            self._selected_forms_set.add(self.entity_listbox.get(i))
            self.entity_listbox.selection_set(i)
        self._refresh_selected_patterns_display()

    def _select_all_patterns(self):
        self._selected_forms_set = set(self._all_forms)
        self._reload_entity_listbox()

    def _clear_all_patterns(self):
        self._selected_forms_set.clear()
        self.entity_listbox.selection_clear(0, "end")
        self._refresh_selected_patterns_display()

    def _browse_folder(self):
        folder = filedialog.askdirectory(initialdir=self.output_dir_var.get())
        if folder: self.output_dir_var.set(folder)

    def _log(self, msg):
        def append():
            self.log_text.config(state="normal")
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.log_text.config(state="disabled")
        self.after(0, append)

    def _on_find_click(self):
        country = self.country_var.get().strip()
        good_pattern = self.good_id_var.get().strip()
        direct_pattern = self.direct_id_var.get().strip()

        if not good_pattern or not direct_pattern:
            messagebox.showerror("Missing Input", "Good ID Pattern aur Direct ID Pattern dono bharna zaroori hai.")
            return

        if good_pattern == direct_pattern:
            messagebox.showerror("Invalid Input", "Good ID Pattern aur Direct ID Pattern alag hone chahiye.")
            return

        self._sync_visible_selection_to_set()
        selected_forms = sorted(self._selected_forms_set)
        custom_words = [w for w in self.custom_pattern_var.get().split(",") if w.strip()]

        if not selected_forms and not custom_words:
            if not messagebox.askyesno("No Entity Pattern Selected", "Bina cleaning ke match karna chahte ho?"): return

        entity_regex = patterns_to_regex(selected_forms, custom_words, self.legal_form_df, suffix_map=self.suffix_map)
        output_dir = self.output_dir_var.get().strip() or self.output_dir
        safe_country = re.sub(r"[^A-Za-z0-9_-]", "_", country) if country else "Output"
        output_path = os.path.join(output_dir, f"Matching_{safe_country}.xlsx")

        params = {
            "country": country,
            "good_pattern": good_pattern,
            "direct_pattern": direct_pattern,
            "entity_regex": entity_regex,
            "threshold": self.threshold_var.get(),
            "address_threshold": self.addr_threshold_var.get(),
            "output_path": output_path,
            "use_hierarchy": self.use_hierarchy_var.get(),
            "suffix_map": self.suffix_map,
            "mandatory_words": self.mandatory_words,
            "geo_rules": self.compiled_geo_rules,          # [NEW] Passing Geo Rules
            "dynamic_geo_regex": self.dynamic_geo_regex    # [NEW] Passing File 6 Regex
        }

        self.find_btn.config(state="disabled")
        self.progress.start(10)
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

        threading.Thread(target=self._run_pipeline_thread, args=(params,), daemon=True).start()

    def _run_pipeline_thread(self, params):
        try:
            output_path = run_matching_pipeline(params, self._log)
            self.after(0, lambda: self._on_success(output_path))
        except Exception as e:
            err_text = f"{e}\n\n{traceback.format_exc()}"
            self.after(0, lambda: self._on_error(err_text))

    def _on_success(self, output_path):
        self.progress.stop()
        self.find_btn.config(state="normal")
        if messagebox.askyesno("Completed", f"Matching complete!\nSaved to:\n{output_path}\n\nOutput folder open karein?"):
            try: os.startfile(os.path.dirname(output_path))
            except Exception: pass

    def _on_error(self, err_text):
        self.progress.stop()
        self.find_btn.config(state="normal")
        self._log(f"[ERROR] {err_text}")
        messagebox.showerror("Error", err_text[:1500])


if __name__ == "__main__":
    app = MatcherApp()
    app.mainloop()