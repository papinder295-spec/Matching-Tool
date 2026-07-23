import tkinter as tk
from tkinter import ttk

def main():
    # Main Window Setup
    root = tk.Tk()
    root.title("Legal Entity Matcher - Good ID / Direct ID Matching Tool")
    root.geometry("780x680")
    root.resizable(False, False)

    pad = {"padx": 10, "pady": 6}

    # ================= TOP SECTION: Inputs =================
    frm_top = ttk.Frame(root)
    frm_top.pack(fill="x", **pad)

    ttk.Label(frm_top, text="Country Name:").grid(row=0, column=0, sticky="w")
    ttk.Entry(frm_top, width=25).grid(row=0, column=1, sticky="w", padx=5)

    ttk.Label(frm_top, text="Good ID Pattern:").grid(row=1, column=0, sticky="w", pady=4)
    ttk.Entry(frm_top, width=25).grid(row=1, column=1, sticky="w", padx=5)
    ttk.Label(frm_top, text="e.g. FIC").grid(row=1, column=2, sticky="w")

    ttk.Label(frm_top, text="Direct ID Pattern:").grid(row=2, column=0, sticky="w", pady=4)
    ttk.Entry(frm_top, width=25).grid(row=2, column=1, sticky="w", padx=5)
    ttk.Label(frm_top, text="e.g. FIC*").grid(row=2, column=2, sticky="w")

    ttk.Label(frm_top, text="Fuzzy Threshold (Name):").grid(row=3, column=0, sticky="w", pady=4)
    ttk.Spinbox(frm_top, from_=50, to=100, width=5).grid(row=3, column=1, sticky="w", padx=5)

    ttk.Label(frm_top, text="Fuzzy Threshold (Address):").grid(row=4, column=0, sticky="w", pady=4)
    ttk.Spinbox(frm_top, from_=30, to=100, width=5).grid(row=4, column=1, sticky="w", padx=5)

    ttk.Checkbutton(frm_top, text="Use hierarchy (links table) validation").grid(row=5, column=0, columnspan=2, sticky="w", pady=4)

    # ================= MIDDLE SECTION: Entity Patterns =================
    frm_entity = ttk.LabelFrame(root, text="Entity Pattern (Legal Form) - Multi Select")
    frm_entity.pack(fill="x", **pad)

    frm_search = ttk.Frame(frm_entity)
    frm_search.pack(fill="x", padx=8, pady=(8, 2))
    ttk.Label(frm_search, text="Search:").pack(side="left")
    ttk.Entry(frm_search, width=30).pack(side="left", padx=5)

    ttk.Button(frm_search, text="Select Visible").pack(side="left", padx=(10, 3))
    ttk.Button(frm_search, text="Select All").pack(side="left", padx=3)
    ttk.Button(frm_search, text="Clear").pack(side="left", padx=3)

    # Listbox for selection
    listbox = tk.Listbox(frm_entity, selectmode="extended", height=7)
    sample_forms = ["AB", "AG", "B.V.", "CORP", "GMBH", "INC", "LLC", "LTD", "OY", "PLC", "PTE", "PVT", "SA", "SPA", "SRL"]
    for form in sample_forms:
        listbox.insert("end", form)
    listbox.pack(side="left", fill="both", expand=True, padx=8, pady=8)

    # Custom pattern inputs
    frm_custom = ttk.Frame(frm_entity)
    frm_custom.pack(side="left", fill="both", expand=True, padx=8, pady=8)
    ttk.Label(frm_custom, text="Custom Patterns\n(comma separated)").pack(anchor="w")
    ttk.Entry(frm_custom, width=35).pack(fill="x", pady=6)
    
    ttk.Label(frm_custom, text="Selected Patterns:").pack(anchor="w", pady=(10,0))
    ttk.Entry(frm_custom, width=35, state="readonly").pack(fill="x", pady=4)

    # ================= BOTTOM SECTION: Output & Run =================
    frm_out = ttk.Frame(root)
    frm_out.pack(fill="x", **pad)
    ttk.Label(frm_out, text="Output Folder:").pack(side="left")
    ttk.Entry(frm_out, width=55).pack(side="left", padx=5)
    ttk.Button(frm_out, text="Browse...").pack(side="left")

    # Run Button
    ttk.Button(root, text="Find Matching").pack(pady=10)
    
    # Progress Bar (dummy)
    ttk.Progressbar(root, mode="indeterminate").pack(fill="x", padx=10)

    # Log Window
    frm_log = ttk.LabelFrame(root, text="Log")
    frm_log.pack(fill="both", expand=True, **pad)
    log_text = tk.Text(frm_log, height=10)
    log_text.insert("1.0", "System Initialized...\nLegal Form Database loaded: 450 entries\nStatic Geo Standardization rules loaded: 14 rules\nDynamic Geo Regex loaded successfully from File 4.\nReady to run!")
    log_text.config(state="disabled") # Make it read-only
    log_text.pack(fill="both", expand=True, padx=5, pady=5)

    # Start App
    root.mainloop()

if __name__ == "__main__":
    main()