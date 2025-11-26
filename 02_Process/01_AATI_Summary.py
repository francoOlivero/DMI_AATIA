import os
import re
from collections import defaultdict

import numpy as np
import pandas as pd

# ---------------------- CONFIG ----------------------
# Get the directory where the Python script lives
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Move one folder up (to DMI_AATIA)
ROOT_DIR = os.path.dirname(BASE_DIR)

# Build paths
INPUT_DIR = os.path.join(ROOT_DIR, "01_Input")
OUTPUT_DIR = os.path.join(ROOT_DIR, "02_Process")  # optional, if saving here

INPUT_CSV  = os.path.join(INPUT_DIR, "ATI_Regular_Life.csv")
OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "Tables_By_Section.xlsx")
# ----------------------------------------------------

def sanitize_sheet_name(name: str) -> str:
    # Excel sheet name rules: max 31 chars; cannot contain : \ / ? * [ ]
    # Also avoid leading/tailing apostrophes.
    if name is None:
        name = "Sheet"
    name = str(name)
    name = re.sub(r'[:\\/\?\*\[\]]', '_', name)
    name = name.replace("'", "’")
    return name[:31] if len(name) > 31 else name

# short module names for sheet naming
MODULE_SHORT = {
    "Regular Life": "RL",
    "Universal Life": "UL",
    "Disability": "DI",
    "Annuities": "AN"
}

def build_sheet_name(module, tbl_type, usage_type, tbl_name, section, shape):
    module_short = MODULE_SHORT.get(module, module)
    base = f"{module_short} | {section} | {tbl_name}"
    return sanitize_sheet_name(base)

def ensure_columns(df: pd.DataFrame, cols):
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise ValueError(f"Missing expected columns in CSV: {missing}")

def coerce_row_to_int(series: pd.Series) -> pd.Series:
    # safe int conversion; non-numeric becomes NaN
    return pd.to_numeric(series, errors='coerce').astype('Int64')

def first_non_null_unique(values):
    vals = [v for v in pd.Series(values).dropna().unique().tolist() if str(v).strip() != ""]
    return vals[0] if vals else None

def concat_unique(values, sep=", "):
    vals = [str(v) for v in pd.Series(values).dropna().unique().tolist() if str(v).strip() != ""]
    return sep.join(vals) if vals else ""

def main():
    # Read CSV: keep raw as string to avoid silent coercions; we’ll explicitly convert selected cols.
    df = pd.read_csv(INPUT_CSV, dtype=str).fillna("") # le tuve que agregar encoding='latin1' para que ande en mi compu, habría que ver dónde va a pasar esto
    # Trim whitespace
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x) # cambiar applymap por solo map

    # Identify C# columns
    c_cols = [c for c in df.columns if re.fullmatch(r"C\d{1,3}", c)]
    # We expect up to C121; keep only those that exist
    c_cols = [c for c in c_cols if int(c[1:]) <= 121]
    # Required base columns
    base_cols = [
        "Section", "Shape", "TableName", "Row", "Op", "LnkSection", "LnkTable",
        "FORMULA", "Name", "Obj Name", "Used By", "Module", "Table Type", "Table_Usage_Type"
    ]

    ensure_columns(df, base_cols) # check required columns

    # Coerce Row to Int64 (nullable)
    df["Row"] = coerce_row_to_int(df["Row"])

    # Build ordering keys (None -> empty)
    for col in ["Module", "Table Type", "Table_Usage_Type", "TableName", "Section", "Shape"]:
        df[col] = df[col].fillna("").astype(str)

    # Sort for stable sheet order: Module, Table Type, Usage, TableName, Row
    sort_cols = ["Module", "Table Type", "Table_Usage_Type", "TableName", "Section", "Shape", "Row"]
    df_sorted = df.sort_values(sort_cols, na_position="last").reset_index(drop=True)

    # Group by the hierarchy for sheet creation (one sheet per TableName)
    groups = df_sorted.groupby(["Module", "Table Type", "Table_Usage_Type", "TableName", "Section", "Shape"], dropna=False)
    print(groups)

    # Track used sheet names to avoid duplicates after sanitization/truncation
    name_counts = defaultdict(int)

    # Prepare an Index dataframe to hyperlink into each sheet
    index_rows = []

    with pd.ExcelWriter(OUTPUT_XLSX, engine="xlsxwriter") as xw:
        wb = xw.book

        # Create Index sheet first
        idx_df = pd.DataFrame(columns=["Module", "Table Type", "Table_Usage_Type", "TableName", "Section", "Shape", "Link"])
        idx_sheet = "Index"
        idx_sheet_safe = sanitize_sheet_name(idx_sheet)
        idx_df.to_excel(xw, sheet_name=idx_sheet_safe, index=False, startrow=0)

        # We'll add hyperlinks after sheets are created
        row_in_index = 1  # 0-based writing with header at row 0 ⇒ first data row = 1

        for (module, tbl_type, usage_type, tbl_name, section, shape), sub in groups:
            # Skip empty TableName (if any)
            if not str(tbl_name).strip():
                continue

            # skip age distribution tables
            if isinstance(tbl_type, str) and tbl_type.strip().lower() == "age distribution":
                print(f"Skipping Age Distribution table: {tbl_name}")
                continue
            
            # delete duplicates
            sub = sub.applymap(lambda x: x.strip() if isinstance(x, str) else x)
            # drop duplicates in Row column
            sub = sub.drop_duplicates(
                subset=[c for c in sub.columns if c not in ["Used By", "Obj Name"]]
            ).copy()

            # Compose desired sheet name and ensure uniqueness within 31 chars
            raw_sheet = build_sheet_name(module, tbl_type, usage_type, tbl_name, section, shape)
            sheet_name = raw_sheet
            if name_counts[sheet_name] > 0:
                # Append a counter while respecting 31-char limit
                suffix = f"_{name_counts[sheet_name]+1}"
                sheet_name = sanitize_sheet_name((raw_sheet[:31 - len(suffix)]) + suffix)
            name_counts[raw_sheet] += 1

            # HEADER BLOCK (meta)
            section_val   = concat_unique(sub["Section"])
            shape_vals    = concat_unique(sub["Shape"])
            used_by_val   = first_non_null_unique(sub["Used By"])
            obj_names_all = concat_unique(sub["Obj Name"])

            header_rows = [
                ["Module", module],
                ["Table Type", tbl_type],
                ["Table_Usage_Type", usage_type],
                ["Section(s)", section_val],
                ["Shape(s)", shape_vals],
                ["Table Name", tbl_name],
                ["Used By", used_by_val],
                ["Obj Name(s)", obj_names_all],
            ]
            header_df = pd.DataFrame(header_rows, columns=["Field", "Value"])

            # NEGATIVE ROW SETTINGS: Row < 0 ⇒ show Row, C1 (value), Op (description)
            neg = sub[(sub["Row"].notna()) & (sub["Row"] < 0)].copy()
            neg_cols = ["Row", "C1", "Op"]
            for col in neg_cols:
                if col not in neg.columns:
                    neg[col] = ""
            neg_settings_df = neg[neg_cols].reset_index(drop=True)
            neg_settings_df.columns = ["Row", "C1 (Setting Value)", "Op (Description)"]

            # FORMULA block: Shape == "S_Rule" and Row == 1
            formula_df = pd.DataFrame(columns=["FORMULA"])
            mask_formula = (sub["Shape"].str.upper() == "S_RULE") & (sub["Row"] == 1)
            if "FORMULA" in sub.columns and mask_formula.any():
                formulas = sub.loc[mask_formula, "FORMULA"].dropna().unique().tolist()
                if formulas:
                    formula_df = pd.DataFrame({"FORMULA": formulas})

            # ASSUMPTION TABLE: Row >= 1 with C# columns
            tbl = sub[(sub["Row"].notna()) & (sub["Row"] >= 1)].copy()
            table_cols = ["Row"] + c_cols
            # Add any missing C# columns as empty to keep consistent shape
            for c in c_cols:
                if c not in tbl.columns:
                    tbl[c] = np.nan
            # Keep only until the last non-all-empty C# column for this TableName to avoid writing tons of empty columns
            # Determine used C# columns
            used_c_cols = []
            for c in c_cols:
                colvals = tbl[c]
                if not (colvals.replace("", np.nan)).isna().all():
                    used_c_cols.append(c)
            if not used_c_cols:
                used_c_cols = []  # If none used, don't write them
            table_final_cols = ["Row"] + used_c_cols
            assumption_df = tbl[table_final_cols].sort_values("Row").reset_index(drop=True)

            # ---- WRITE SHEET ----
            header_start = 0
            neg_start    = header_start + len(header_df) + 2
            formula_start= neg_start + (len(neg_settings_df) + 2 if len(neg_settings_df) > 0 else 2)
            table_start  = formula_start + (len(formula_df) + 2 if len(formula_df) > 0 else 2)

            # Create the sheet by writing header_df first
            header_df.to_excel(xw, sheet_name=sheet_name, index=False, startrow=header_start)
            ws = xw.sheets[sheet_name]

            # Label blocks
            ws.write(neg_start - 1, 0, "Negative Row Settings (Row < 0)")
            if len(neg_settings_df) > 0:
                neg_settings_df.to_excel(xw, sheet_name=sheet_name, index=False, startrow=neg_start)
            else:
                ws.write(neg_start, 0, "(none)")

            ws.write(formula_start - 1, 0, "FORMULA (S_Rule & Row = 1)")
            if len(formula_df) > 0:
                formula_df.to_excel(xw, sheet_name=sheet_name, index=False, startrow=formula_start)
            else:
                ws.write(formula_start, 0, "(none)")

            ws.write(table_start - 1, 0, "Assumption Table (Rows ≥ 1)")
            if len(assumption_df) > 0:
                assumption_df.to_excel(xw, sheet_name=sheet_name, index=False, startrow=table_start)
            else:
                ws.write(table_start, 0, "(none)")

            # Some light formatting: bold for first column of header
            bold = wb.add_format({"bold": True})
            ws.set_column(0, 0, 24, bold)      # Field / first column
            ws.set_column(1, 1, 60)            # Value / second column
            ws.set_zoom(120)

            # Add entry for Index sheet (hyperlink)
            index_rows.append({
                "Module": module,
                "Table Type": tbl_type,
                "Table_Usage_Type": usage_type,
                "TableName": tbl_name,
                "Section": section,
                "Shape": shape,
                "Sheet": sheet_name  # keep for hyperlink writing
            })

        # Write hyperlinks into Index sheet
        if index_rows:
            idx = pd.DataFrame(index_rows).sort_values(["Module", "Table Type", "Table_Usage_Type", "TableName", "Section", "Shape"]).reset_index(drop=True)
            # Rewrite the table (excluding the helper "Sheet" column)
            printable = idx[["Module", "Table Type", "Table_Usage_Type", "TableName", "Section", "Shape"]]
            printable.to_excel(xw, sheet_name=idx_sheet_safe, index=False, startrow=0)
            idx_ws = xw.sheets[idx_sheet_safe]

            # Add hyperlinks in a 5th column "Link"
            idx_ws.write(0, 6, "Link")
            for i, row in idx.iterrows():
                # Hyperlink to A1 in the target sheet
                idx_ws.write_url(i + 1, 6, f"internal:'{row['Sheet']}'!A1", string="Open")
            idx_ws.set_column(0, 3, 28)
            idx_ws.set_column(4, 4, 10)
            idx_ws.freeze_panes(1, 0)

    print(f"Done. Wrote: {OUTPUT_XLSX}")

if __name__ == "__main__":
    main()