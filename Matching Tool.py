"""
=====================================================================================
LEGAL ENTITY MATCHER - DESKTOP TOOL  (Country-agnostic, GUI based)
=====================================================================================
Ye ek Tkinter based desktop application hai jo company laptop pr directly chalegi
(Python + ODBC driver installed hone chahiye). Isme UI se hi:

    - Country Name          (sirf label / output file naming k liye)
    - Good ID Pattern        (e.g. "FIC")
    - Direct ID Pattern      (e.g. "FIC*")
    - Entity / Legal-Form Patterns   (multi-select list + custom entry)
        -> Ye woh legal form suffixes hain jo company name se clean/remove honge
           taaki fuzzy matching sahi ho (jaise Finland ke liye: OY, OYK, LTD, LIMITED)
    - Fuzzy Threshold
    - "Find Matching" button dabate hi ek Excel file ban kar output folder me save
      ho jayegi.

-------------------------------------------------------------------------------------
FIXED RULES (in par tool khud dhyan deta hai, UI se change nahi hote):
-------------------------------------------------------------------------------------
  1. Good ID aur Direct ID hamesha ALAG record hone chahiye (khud check hota hai)
  2. (Person NOT IN ('I','L','K','D') OR Person IS NULL)
  3. (Branch IS NULL OR Branch IN (0, 3))
  4. ([Foreign] IS NULL OR [Foreign] IN (0, 3))
  5. Company Name (legal-form clean karne ke baad) match hona chahiye
  6. Agar dono (Direct + Good) ke paas Address/City hai, to Address bhi match
     hona chahiye. Agar Direct ke paas Address/City hi nahi hai (blank/NULL),
     to sirf Name match ke basis pr match maan liya jayega (Address requirement
     tab skip ho jata hai).
  7. Fuzzy Score (Name aur agar available ho to Address) output me hamesha diya
     jayega.
  8. Agar ek Direct ID ke against Name-match se MULTIPLE Good IDs aa rahe hain
     aur Hierarchy (links table) se bhi disambiguate nahi ho pa raha, to
     Address/PCCity ko tie-breaker ki tarah use kiya jayega: agar in multiple
     candidates me se SIRF EK ka Address/PCCity bhi match karta hai, to usi ko
     final match maan liya jayega ("Identified by Address/City Match").

-------------------------------------------------------------------------------------
DATA FILES (same folder me rakhna hai):
-------------------------------------------------------------------------------------
  - legal_form_mapping.xlsx        -> Sheet "in"  : LegalForm, Regex, RelatedLegalForms
  - Universal_Suffix_Map.xlsx.xlsx -> Sheet "Sheet3": LegalForm, Regex, RelatedLegalForms
                                   -> Sheet "Suffix_Map": Suffix, Replacement, Country
                                   -> Sheet "Mandatory_Words": Word

-------------------------------------------------------------------------------------
SETUP (company laptop pr ek baar karna hai):
-------------------------------------------------------------------------------------
  pip install pandas sqlalchemy pyodbc thefuzz python-Levenshtein openpyxl rapidfuzz

  Neeche CONFIG section me apna SQL Server / Database / column names check kar lena
  (Address / City column ka naam database schema k hisaab se alag ho sakta hai).
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

# rapidfuzz process.cdist -> ek hi call me pura batch (list vs list) fuzzy score
# nikal deta hai (C-optimized), row-by-row .apply() se kaafi tez hai.
try:
    from rapidfuzz import process as rf_process, fuzz as rf_fuzz
except ImportError:
    rf_process = None
    rf_fuzz = None

import tkinter as tk
from tkinter import ttk, messagebox, filedialog


# =====================================================================================
# 1. CONFIGURATION -- DATABASE + SCHEMA (company laptop k hisaab se yaha adjust karna)
# =====================================================================================

DB_CONFIG = {
    "SQL_SERVER": "QIG-WXRELADB501.analytics.moodys.net",
    "DATABASE": "bvdaffils",
    "ODBC_DRIVER": "ODBC Driver 17 for SQL Server",
}

# Companies table me Address/City columns ka actual naam yaha set karo
# (agar database me ye columns kisi aur naam se hain to bas yaha update kar dena)
SCHEMA_CONFIG = {
    "ID_COL": "Id",
    "NAME_COL": "Name",
    "ADDRESS_COL": "Address",     # <-- adjust if different in your schema
    "CITY_COL": "PCCity",         # <-- adjust if different in your schema
    "NRLINKS_COL": "NrLinks",
    "BRANCH_COL": "Branch",
    "FOREIGN_COL": "[Foreign]",
    "PERSON_COL": "Person",
    "SOURCENR_COL": "SourceNr",
    "TABLE_NAME": "Companies",
    "LINKS_TABLE": "links",
}

# ---- Script ki apni directory (Excel files yahi se load hongi) ----
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ---- Legal Form / Suffix Data Sources ----
# Ye 2 files se saara "Entity Pattern" data (Listbox), global suffix normalization,
# aur Mandatory Words list load hoti hai. In files ko is script k saath (same folder)
# rakhna hai, ya neeche full path de sakte ho.
LEGAL_FORM_SOURCES = [
    # (file_path, sheet_name)
    (os.path.join(SCRIPT_DIR, "legal_form_mapping.xlsx"), "in"),
    (os.path.join(SCRIPT_DIR, "Universal_Suffix_Map.xlsx.xlsx"), "Sheet3"),
]
SUFFIX_MAP_SOURCE = (os.path.join(SCRIPT_DIR, "Universal_Suffix_Map.xlsx.xlsx"), "Suffix_Map")
MANDATORY_WORDS_SOURCE = (os.path.join(SCRIPT_DIR, "Universal_Suffix_Map.xlsx.xlsx"), "Mandatory_Words")

DEFAULT_THRESHOLD = 88
DEFAULT_ADDRESS_THRESHOLD = 70


# =====================================================================================
# 2. LEGAL FORM DATABASE (Entity Pattern list ke liye)
# =====================================================================================

def _load_single_legal_form_sheet(file_path, sheet_name):
    """
    Ek Excel file ki ek specific sheet se legal-form data padh ke DataFrame return
    karta hai with columns ['LegalForm', 'Regex'].
    """
    if not os.path.exists(file_path):
        print(f"[WARNING] Legal form file not found: {file_path}")
        return pd.DataFrame(columns=["LegalForm", "Regex"])

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        print(f"[WARNING] Cannot read sheet '{sheet_name}' from {file_path}: {e}")
        return pd.DataFrame(columns=["LegalForm", "Regex"])

    df.columns = [str(c).strip() for c in df.columns]

    # Auto-detect columns by name
    name_col = next(
        (c for c in df.columns if "legal" in c.lower() and "form" in c.lower()),
        next((c for c in df.columns if c.lower() == "legalform"), df.columns[0])
    )
    regex_col = next((c for c in df.columns if "regex" in c.lower()), None)

    out = pd.DataFrame()
    out["LegalForm"] = df[name_col].astype(str).str.strip()
    out["Regex"] = df[regex_col].astype(str).str.strip() if regex_col else ""
    out = out[out["LegalForm"].notna() & (out["LegalForm"] != "") & (out["LegalForm"] != "nan")]
    return out


def load_legal_form_database(sources=None):
    """
    Multiple Excel files + sheets se legal-form data merge kar ke ek single DataFrame
    return karta hai. Duplicates hata deta hai (LegalForm column ke basis pe).

    Parameters:
        sources: list of (file_path, sheet_name) tuples. None hone par LEGAL_FORM_SOURCES use hoga.

    Returns:
        DataFrame with columns ['LegalForm', 'Regex'], duplicates removed, sorted.
    """
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
    # Duplicate removal: pehle wala rakhte hain (legal_form_mapping.xlsx ki priority)
    merged = merged.drop_duplicates(subset=["LegalForm"], keep="first")
    merged = merged.sort_values("LegalForm").reset_index(drop=True)
    return merged


def load_suffix_map(source=None):
    """
    Suffix_Map sheet se abbreviation → full form normalization map load karta hai.
    Example: "ltd" → "limited", "inc" → "incorporated", "corp" → "corporation"

    Ye name cleaning me USE hota hai: pehle suffix normalize hoga, phir legal form
    regex hatayega. Isse accuracy improve hoti hai kyunki "ABC LTD" aur "ABC LIMITED"
    dono same clean name denge.

    Returns:
        dict: {suffix_lowercase: replacement_lowercase} e.g. {"ltd": "limited", ...}
    """
    if source is None:
        source = SUFFIX_MAP_SOURCE

    file_path, sheet_name = source
    if not os.path.exists(file_path):
        print(f"[WARNING] Suffix map file not found: {file_path}")
        return {}

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        print(f"[WARNING] Cannot read Suffix_Map: {e}")
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
        # Skip header-like rows or NaN
        if suffix in ("", "nan", "common/global") or replacement in ("", "nan"):
            continue
        suffix_map[suffix] = replacement

    return suffix_map


def load_mandatory_words(source=None):
    """
    Mandatory_Words sheet se protected words load karta hai.
    Ye words company name me important hain aur cleaning process me PRESERVE hone
    chahiye (hata NAHI dena).

    Example: "holding", "group", "invest" — agar company name "ABC HOLDING LTD" hai
    to "HOLDING" bachna chahiye, sirf "LTD" hatna chahiye.

    Returns:
        set: lowercase mandatory words e.g. {"holding", "group", "invest", ...}
    """
    if source is None:
        source = MANDATORY_WORDS_SOURCE

    file_path, sheet_name = source
    if not os.path.exists(file_path):
        print(f"[WARNING] Mandatory words file not found: {file_path}")
        return set()

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        print(f"[WARNING] Cannot read Mandatory_Words: {e}")
        return set()

    df.columns = [str(c).strip() for c in df.columns]
    word_col = df.columns[0]  # First column has the words

    words = set()
    for val in df[word_col].dropna():
        w = str(val).strip().lower()
        if w and w != "nan":
            words.add(w)

    return words


def patterns_to_regex(selected_forms, custom_words, legal_form_df):
    """
    User ne jo Entity Patterns select/type kiye hain, unse ek combined regex banata hai
    jo company name se legal-form suffix hata degi. Agar select kiya hua form mapping
    table me mila (uska proper regex mil gaya) to wahi use hoga, warna simple
    \\bWORD\\b regex ban jayega.
    """
    parts = []

    for form in selected_forms:
        row = legal_form_df[legal_form_df["LegalForm"].str.upper() == str(form).upper()]
        if not row.empty and str(row.iloc[0]["Regex"]).strip() not in ("", "nan"):
            pattern = row.iloc[0]["Regex"]
            try:
                re.compile(pattern)
                parts.append(f"(?:{pattern})")
                continue
            except re.error:
                pass
        # fallback: simple word-boundary pattern
        escaped = re.escape(str(form).strip())
        parts.append(rf"\b{escaped}\b")

    for word in custom_words:
        word = word.strip()
        if not word:
            continue
        escaped = re.escape(word)
        parts.append(rf"\b{escaped}\b")

    if not parts:
        return None
    return re.compile("|".join(parts), flags=re.IGNORECASE)


# =====================================================================================
# 3. CLEANING FUNCTIONS
# =====================================================================================

def make_clean_name_fn(entity_regex, suffix_map=None, mandatory_words=None):
    """
    Name cleaning function factory.

    Cleaning pipeline:
      1. Lowercase
      2. Extract mandatory words present in original name (for protection)
      3. Legal Form Removal: entity_regex se legal suffixes hata deta hai
         (e.g. LTD, GMBH, OY etc. user ki selection ke hisaab se)
      4. Special characters removal + whitespace normalization
      5. Suffix Normalization: abbreviations → full forms (ltd→limited, inc→incorporated)
         Ye REMAINING name pe apply hota hai taaki "ABC LTD" aur "ABC LIMITED"
         dono same clean name denge (agar LTD selected nahi tha to)
      6. Mandatory Words Protection: agar regex ne koi mandatory word hata diya
         (like "holding", "group") to wapas add kar deta hai

    Parameters:
        entity_regex: compiled regex for removing legal form suffixes
        suffix_map: dict {abbreviation: full_form} for normalization
        mandatory_words: set of words that must be preserved in the name
    """
    if suffix_map is None:
        suffix_map = {}
    if mandatory_words is None:
        mandatory_words = set()

    # Pre-compile suffix normalization regex for performance
    if suffix_map:
        # Sort by length descending so longer suffixes match first (e.g. "co ltd" before "ltd")
        sorted_suffixes = sorted(suffix_map.keys(), key=len, reverse=True)
        suffix_patterns = [rf"\b{re.escape(s)}\b" for s in sorted_suffixes]
        suffix_regex = re.compile("|".join(suffix_patterns), flags=re.IGNORECASE)
    else:
        suffix_regex = None

    @lru_cache(maxsize=None)
    def clean_name(name):
        if pd.isna(name) or not name:
            return ""
        cleaned = str(name).lower().strip()

        # Step 1: Extract mandatory words present BEFORE any removal
        preserved_words = set()
        if mandatory_words:
            name_words = set(re.findall(r'[a-z]+', cleaned))
            preserved_words = name_words & mandatory_words

        # Step 2: Remove legal form suffixes via entity regex
        #         (e.g. LTD, GMBH, OY etc. hatenge pehle)
        if entity_regex is not None:
            cleaned = entity_regex.sub(" ", cleaned)
            # Pattern hatane ke turant baad jo extra/double space reh jata hai
            # usko yahin par clean kar dete hain (isse baad ke steps me bhi
            # koi stray space carry forward nahi hoga)
            cleaned = re.sub(r"\s+", " ", cleaned).strip()

        # Step 3: Remove special characters, keep only alphanumeric + spaces
        cleaned = re.sub(r"[^a-z0-9\s]", " ", cleaned)
        cleaned = re.sub(r"\s+", " ", cleaned).strip()

        # Step 4: Suffix Normalization on REMAINING name
        #         (ltd→limited, inc→incorporated etc.)
        #         Ye tab useful hai jab user ne specific legal form select nahi kiya
        #         lekin dono sides me abbreviation/full form ka mismatch hai
        if suffix_regex:
            def replace_suffix(match):
                matched_text = match.group(0).lower()
                return suffix_map.get(matched_text, matched_text)
            cleaned = suffix_regex.sub(replace_suffix, cleaned)
            cleaned = re.sub(r"\s+", " ", cleaned).strip()

        # Step 5: Restore mandatory words if they were accidentally removed
        if preserved_words:
            current_words = set(cleaned.split())
            missing_words = preserved_words - current_words
            if missing_words:
                cleaned = cleaned + " " + " ".join(sorted(missing_words))
                cleaned = re.sub(r"\s+", " ", cleaned).strip()

        return cleaned
    return clean_name


@lru_cache(maxsize=None)
def clean_address(addr):
    if pd.isna(addr) or not addr or str(addr).strip() == "":
        return ""
    cleaned = str(addr).lower()
    cleaned = re.sub(r"[^a-z0-9\s]", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


# =====================================================================================
# 4. SQL QUERY BUILDERS
# =====================================================================================

def build_company_query(good_pattern, direct_pattern, schema):
    """
    good_pattern e.g. "FIC"  -> Good Id = LIKE 'FIC%' AND NOT LIKE 'FIC*%' (agar direct me * hai)
    direct_pattern e.g. "FIC*" -> Direct Id = LIKE 'FIC*%'
    """
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

    select_cols = f"{id_col}, {name_col}, {addr_col}, {city_col}, {nrlinks_col}, {branch_col}, {foreign_col}, {sourcenr_col}"

    fixed_filters = f"""
      AND ({person_col} NOT IN ('I', 'L', 'K', 'D') OR {person_col} IS NULL)
      AND ({branch_col} IS NULL OR {branch_col} IN (0, 3))
      AND ({foreign_col} IS NULL OR {foreign_col} IN (0, 3))
    """

    query = f"""
WITH GoodCompanies AS (
    SELECT {select_cols}
    FROM {table}
    WHERE
    ({id_col} LIKE '{good_pattern}%' AND {id_col} NOT LIKE '{direct_pattern.rstrip('%')}%')
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
# 5. MATCHING PIPELINE (runs in background thread, log_fn se GUI ko status milta hai)
# =====================================================================================

def run_matching_pipeline(params, log_fn):
    """
    params dict keys:
        country, good_pattern, direct_pattern, entity_regex,
        threshold, address_threshold, output_path, use_hierarchy,
        suffix_map, mandatory_words
    """
    if create_engine is None or fuzz is None:
        raise RuntimeError(
            "Required libraries missing. Company laptop pr chalao:\n"
            "pip install sqlalchemy pyodbc thefuzz python-Levenshtein"
        )

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
    try:
        df = pd.read_sql(company_query, engine)
    except Exception as e:
        raise RuntimeError(
            f"Company data fetch fail hua. Agar error 'Invalid column name' bol raha hai "
            f"to SCHEMA_CONFIG me ADDRESS_COL / CITY_COL check kar lena.\n\nDetails: {e}"
        )

    use_hierarchy = params.get("use_hierarchy", True)
    parent_map, child_map = {}, {}
    if use_hierarchy:
        log_fn("Fetching hierarchy (links) data...")
        try:
            links_df = pd.read_sql(links_query, engine)
            parent_map = links_df.groupby("ParentId")["ChildId"].apply(set).to_dict()
            child_map = links_df.groupby("ChildId")["ParentId"].apply(set).to_dict()
        except Exception as e:
            log_fn(f"[WARNING] Links table fetch fail hui, hierarchy validation skip: {e}")

    id_col = schema["ID_COL"]
    name_col = schema["NAME_COL"]
    addr_col = schema["ADDRESS_COL"]
    city_col = schema["CITY_COL"]
    nrlinks_col = schema["NRLINKS_COL"]
    sourcenr_col = schema["SOURCENR_COL"]

    # Use enhanced clean_name with suffix normalization + mandatory words protection
    clean_name = make_clean_name_fn(
        params["entity_regex"],
        suffix_map=params.get("suffix_map"),
        mandatory_words=params.get("mandatory_words"),
    )

    log_fn("Cleaning names & addresses...")
    df["Cleaned Name"] = df[name_col].map(clean_name)
    df["Cleaned Address"] = (df[addr_col].fillna("") + " " + df[city_col].fillna("")).map(clean_address)
    df["Has_Address"] = df["Cleaned Address"].str.strip() != ""

    good_df = df[df["Source"] == "Good"].copy()
    direct_df = df[df["Source"] == "Direct"].copy()

    # ---- BLOCKING / GROUPING (Performance ka sabse important step) ----
    # Har Direct ID ko PURE Good ID pool (e.g. 10 lakh records) ke against check
    # karne ke bajaye, hum Good IDs ko unke Cleaned Name ke PEHLE WORD se groups
    # (buckets) me bant dete hain -> "ABC Holding Limited", "ABC Trade Limited",
    # "ABC Corp" -> teeno ka First_Word "abc" hai, to teeno EK hi group me aa jate
    # hain. Ab jab "ABC Ltd" wali Direct ID aati hai, uska bhi First_Word "abc"
    # nikaal ke SEEDHA usi chhote group (na ki poore 10 lakh records) me fuzzy
    # match dhoonda jata hai. Isse effective comparison O(N x M) na hoke
    # O(N x avg_group_size) ho jata hai - bahut fast.
    good_df["First_Word"] = good_df["Cleaned Name"].apply(lambda x: x.split()[0] if x else None)
    direct_df["First_Word"] = direct_df["Cleaned Name"].apply(lambda x: x.split()[0] if x else None)

    good_groups = good_df.groupby("First_Word")

    threshold = params["threshold"]
    addr_threshold = params["address_threshold"]

    all_matches, no_match_list = [], []
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
                "Direct ID": d_id, "Direct Original Name": d_name, "Direct SourceNr": d_sn
            })
            continue

        candidates = good_groups.get_group(d_fw).copy()
        # Rule: Good ID aur Direct ID alag hone chahiye
        candidates = candidates[candidates[id_col] != d_id]
        if candidates.empty:
            no_match_list.append({
                "Direct ID": d_id, "Direct Original Name": d_name, "Direct SourceNr": d_sn
            })
            continue

        # ---- Name scoring (group ke andar hi, chhota sa candidate set) ----
        # rapidfuzz.process.cdist ek single vectorized/C-level call me pure batch
        # ka score nikal deta hai - row-by-row .apply() se kaafi tez hai, khaaskar
        # jab groups bade ho (e.g. common first words jaise "global", "the" etc.)
        cand_names = candidates["Cleaned Name"].tolist()
        if rf_process is not None:
            candidates["Name_Score"] = rf_process.cdist([c_name], cand_names, scorer=rf_fuzz.ratio)[0]
        else:
            candidates["Name_Score"] = [fuzz.ratio(c_name, x) for x in cand_names]

        # Address rule: agar Direct ke paas address hai AND candidate ke paas bhi hai,
        # to address score bhi threshold pura karna hoga. Direct ke paas address na ho
        # to sirf Name match kaafi hai.
        if has_addr:
            cand_addrs = candidates["Cleaned Address"].tolist()
            cand_has_addr = candidates["Has_Address"].tolist()
            if rf_process is not None:
                raw_addr_scores = rf_process.cdist([c_addr], cand_addrs, scorer=rf_fuzz.ratio)[0]
            else:
                raw_addr_scores = [fuzz.ratio(c_addr, x) for x in cand_addrs]
            candidates["Address_Score"] = [
                score if has else None for score, has in zip(raw_addr_scores, cand_has_addr)
            ]
        else:
            candidates["Address_Score"] = None

        def passes(r):
            if r["Name_Score"] < threshold:
                return False
            if has_addr and r["Has_Address"]:
                return r["Address_Score"] is not None and r["Address_Score"] >= addr_threshold
            return True  # address check skip - Direct ya Good me address missing hai

        found = candidates[candidates.apply(passes, axis=1)]
        if found.empty:
            no_match_list.append({
                "Direct ID": d_id, "Direct Original Name": d_name, "Direct SourceNr": d_sn
            })
            continue

        for _, m in found.iterrows():
            all_matches.append({
                "Direct ID": d_id, "Direct Original Name": d_name, "Direct SourceNr": d_sn,
                "Matched Good ID": m[id_col], "Matched Good Original Name": m[name_col],
                "Matched SourceNr": m[sourcenr_col],
                "Name Fuzzy Score": m["Name_Score"],
                "Address Fuzzy Score": m["Address_Score"] if pd.notna(m["Address_Score"]) else "N/A (no address)",
                "Direct NrLinks": d_nl, "Matched NrLinks": m[nrlinks_col],
                "Direct Cleaned Name": c_name, "Matched Cleaned Name": m["Cleaned Name"],
            })

    log_fn("Consolidating matches (hierarchy validation + best match selection)...")
    all_matches_df = pd.DataFrame(all_matches)
    single_match_list, multi_match_list = [], []

    if not all_matches_df.empty:
        for did, group in all_matches_df.groupby("Direct ID"):
            filtered = apply_hierarchy_filter(group, parent_map, child_map)

            # Agar Hierarchy Link se disambiguate nahi hua (abhi bhi multiple
            # matches hain), to Address/PCCity ko tie-breaker ki tarah try karo
            if len(filtered) > 1:
                filtered = apply_address_filter(filtered, addr_threshold)

            # Agar ek hi match filter hoke aaya (ya toh ek hi tha, ya hierarchy/address ne sirf ek ko chuna)
            if len(filtered) == 1:
                single_match_list.append(filtered.iloc[0].to_dict())
            # Agar abhi bhi multiple matches hain, toh tie-break mat karo, sabko multi-match me daal do
            else:
                multi_match_list.append(filtered)

    single_match_df = pd.DataFrame(single_match_list)
    multi_match_df = pd.concat(multi_match_list) if multi_match_list else pd.DataFrame()

    final_export_cols = [
        "Direct ID", "Direct Original Name", "Direct SourceNr", "Matched Good ID",
        "Matched Good Original Name", "Matched SourceNr", "Name Fuzzy Score",
        "Address Fuzzy Score", "Validation_Message", "Direct NrLinks", "Matched NrLinks",
    ]

    def get_clean_df(df_in):
        return df_in.reindex(columns=final_export_cols)

    if not single_match_df.empty:
        hierarchy_matches = single_match_df[single_match_df["Validation_Message"].str.contains("Link", na=False)]
        address_matches = single_match_df[single_match_df["Validation_Message"].str.contains("Address/City", na=False)]
        fuzzy_matches = single_match_df[
            ~single_match_df["Validation_Message"].str.contains("Link", na=False)
            & ~single_match_df["Validation_Message"].str.contains("Address/City", na=False)
        ]
        hierarchy_out = get_clean_df(hierarchy_matches)
        address_out = get_clean_df(address_matches)
        fuzzy_out = get_clean_df(fuzzy_matches)
        high_score_no_link = get_clean_df(
            single_match_df[(single_match_df["Name Fuzzy Score"] >= 95) & (single_match_df["Direct NrLinks"] == 0)]
        )
    else:
        hierarchy_out = address_out = fuzzy_out = high_score_no_link = pd.DataFrame(columns=final_export_cols)

    multi_match_out = get_clean_df(multi_match_df) if not multi_match_df.empty else pd.DataFrame(columns=final_export_cols)

    no_match_df = pd.DataFrame(no_match_list)

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


def apply_hierarchy_filter(group, parent_map, child_map):
    """
    Hierarchy validation logic:

    SUBSIDIARY CHECK:
      - Agar Direct ID parent_map me hai (matlab Direct kisi ka parent hai)
      - To check karo: Matched Good ID uska child hai?
        YA Matched Good ID ka bhi koi parent same hai? (sibling through parent)
      - Agar haan → "Identified by Subsidiary Link"

    SHAREHOLDER CHECK:
      - Agar Direct ID child_map me hai (matlab Direct kisi ka child hai)
      - To check karo: Matched Good ID uska parent hai?
        YA Matched Good ID ka bhi koi child same hai? (sibling through child)
      - Agar haan → "Identified by Shareholder Link"
    """
    did = group["Direct ID"].iloc[0]
    if did in parent_map:
        matches = group[group["Matched Good ID"].apply(
            lambda x: x in parent_map[did] or bool(parent_map.get(x, set()) & parent_map[did])
        )]
        if not matches.empty:
            return matches.assign(Validation_Message="Identified by Subsidiary Link")
    if did in child_map:
        matches = group[group["Matched Good ID"].apply(
            lambda x: x in child_map[did] or bool(child_map.get(x, set()) & child_map[did])
        )]
        if not matches.empty:
            return matches.assign(Validation_Message="Identified by Shareholder Link")
    
    if len(group) == 1:
        return group.assign(Validation_Message="Single Match (No Hierarchy)")
    else:
        return group.assign(Validation_Message="Multiple Matches (No Hierarchy)")


def apply_address_filter(group, addr_threshold):
    """
    ADDRESS / PCCITY CHECK (Hierarchy Validation jaisa hi ek disambiguation step):

    Agar Name-match ke baad (aur Hierarchy Link se) bhi ek Direct ID ke against
    MULTIPLE Good IDs match kar rahe hain, to Address/PCCity ko tie-breaker
    ki tarah use karo:
      - Un candidates me se sirf wahi count karo jinka Address Fuzzy Score
        addr_threshold se >= hai.
      - Agar aisa SIRF EK hi candidate hai (baaki sab address pe match nahi
        karte), to usi ko sahi match maan lo -> "Identified by Address/City Match"
      - Agar 0 ya 1 se zyada candidates address pe bhi qualify karte hain,
        to ambiguity waise hi rehne do (group ko unchanged return karo) taaki
        wo Multi Match sheet me hi jaye - galat tie-break na ho.
    """
    if len(group) <= 1:
        return group

    addr_scores_numeric = pd.to_numeric(group["Address Fuzzy Score"], errors="coerce")
    qualifying = group[addr_scores_numeric >= addr_threshold]

    if len(qualifying) == 1:
        return qualifying.assign(Validation_Message="Identified by Address/City Match")

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

        # ---- Load all data from Excel files ----
        self.legal_form_df = load_legal_form_database()
        self.suffix_map = load_suffix_map()
        self.mandatory_words = load_mandatory_words()

        self.output_dir = os.path.join(os.path.expanduser("~"), "Desktop")

        self._build_ui()

        # Show load summary in log
        self._log(f"Legal Form Database loaded: {len(self.legal_form_df)} entries "
                  f"(from {len(LEGAL_FORM_SOURCES)} source(s))")
        self._log(f"Suffix Map loaded: {len(self.suffix_map)} normalization rules "
                  f"(e.g. ltd→limited, inc→incorporated)")
        self._log(f"Mandatory Words loaded: {len(self.mandatory_words)} words "
                  f"({', '.join(sorted(self.mandatory_words))})")

    # -------------------------------------------------------------------------
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
        ttk.Checkbutton(
            frm_top, text="Use hierarchy (links table) validation",
            variable=self.use_hierarchy_var
        ).grid(row=5, column=0, columnspan=2, sticky="w", pady=4)

        # ---- Entity Pattern selection ----
        frm_entity = ttk.LabelFrame(self, text="Entity Pattern (Legal Form) - Multi Select")
        frm_entity.pack(fill="x", **pad)

        # Search box for filtering the listbox
        frm_search = ttk.Frame(frm_entity)
        frm_search.pack(fill="x", padx=8, pady=(8, 2))
        ttk.Label(frm_search, text="Search:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._filter_entity_list)
        ttk.Entry(frm_search, textvariable=self.search_var, width=30).pack(side="left", padx=5)
        self._all_forms = sorted(self.legal_form_df["LegalForm"].tolist())

        self.entity_listbox = tk.Listbox(frm_entity, selectmode="extended", height=8, exportselection=False)
        for form in self._all_forms:
            self.entity_listbox.insert("end", form)
        self.entity_listbox.pack(side="left", fill="both", expand=True, padx=8, pady=8)

        scrollbar = ttk.Scrollbar(frm_entity, orient="vertical", command=self.entity_listbox.yview)
        scrollbar.pack(side="left", fill="y")
        self.entity_listbox.config(yscrollcommand=scrollbar.set)

        frm_custom = ttk.Frame(frm_entity)
        frm_custom.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        ttk.Label(frm_custom, text="Custom Patterns\n(comma separated, agar list me\nnahi milein to yaha likh do)\ne.g. OYK, LTD, LIMITED").pack(anchor="w")
        self.custom_pattern_var = tk.StringVar()
        ttk.Entry(frm_custom, textvariable=self.custom_pattern_var, width=35).pack(fill="x", pady=6)

        # ---- Output folder ----
        frm_out = ttk.Frame(self)
        frm_out.pack(fill="x", **pad)
        ttk.Label(frm_out, text="Output Folder:").pack(side="left")
        self.output_dir_var = tk.StringVar(value=self.output_dir)
        ttk.Entry(frm_out, textvariable=self.output_dir_var, width=55).pack(side="left", padx=5)
        ttk.Button(frm_out, text="Browse...", command=self._browse_folder).pack(side="left")

        # ---- Action button ----
        self.find_btn = ttk.Button(self, text="Find Matching", command=self._on_find_click)
        self.find_btn.pack(pady=10)

        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.pack(fill="x", padx=10)

        # ---- Log box ----
        frm_log = ttk.LabelFrame(self, text="Log")
        frm_log.pack(fill="both", expand=True, **pad)
        self.log_text = tk.Text(frm_log, height=12, state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)

    # -------------------------------------------------------------------------
    def _filter_entity_list(self, *args):
        """Search box se entity listbox ko filter karta hai (live search)."""
        search_term = self.search_var.get().strip().lower()
        self.entity_listbox.delete(0, "end")
        for form in self._all_forms:
            if not search_term or search_term in form.lower():
                self.entity_listbox.insert("end", form)

    # -------------------------------------------------------------------------
    def _browse_folder(self):
        folder = filedialog.askdirectory(initialdir=self.output_dir_var.get())
        if folder:
            self.output_dir_var.set(folder)

    def _log(self, msg):
        def append():
            self.log_text.config(state="normal")
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.log_text.config(state="disabled")
        self.after(0, append)

    # -------------------------------------------------------------------------
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

        selected_indices = self.entity_listbox.curselection()
        selected_forms = [self.entity_listbox.get(i) for i in selected_indices]
        custom_words = [w for w in self.custom_pattern_var.get().split(",") if w.strip()]

        if not selected_forms and not custom_words:
            if not messagebox.askyesno(
                "No Entity Pattern Selected",
                "Aapne koi Entity/Legal-Form Pattern select nahi kiya. Bina cleaning ke "
                "match karna chahte ho?"
            ):
                return

        entity_regex = patterns_to_regex(selected_forms, custom_words, self.legal_form_df)

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
        }

        self.find_btn.config(state="disabled")
        self.progress.start(10)
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

        thread = threading.Thread(target=self._run_pipeline_thread, args=(params,), daemon=True)
        thread.start()

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
            folder = os.path.dirname(output_path)
            try:
                os.startfile(folder)  # Windows only
            except Exception:
                pass

    def _on_error(self, err_text):
        self.progress.stop()
        self.find_btn.config(state="normal")
        self._log(f"[ERROR] {err_text}")
        messagebox.showerror("Error", err_text[:1500])


if __name__ == "__main__":
    app = MatcherApp()
    app.mainloop()
