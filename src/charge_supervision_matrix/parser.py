"""
Parse a raw 'All Signed Charges' Excel export into a clean DataFrame.

Two report types are supported:

  inpatient  — "All Signed Charges" export grouped by Supervising MD
               Supervisor column: col 26 ("Supervising MD")
               CPT: col 21 | Qty: col 23 | Signed-Off By: col 31

  outpatient — "All Signed Charges (Quick Run)" export grouped by Order MD
               Supervisor column: col 26 ("Order MD")
               CPT: col 18 | Qty: col 22 | Signed-Off By: col 30
"""

import re
import pandas as pd

_HEADER_ROW_MARKER = "Patient"

# Inpatient column map (grouped by Supervising MD)
_INPATIENT_COL = {
    "patient": 1,
    "mrn": 3,
    "description": 4,
    "charge_type": 18,
    "cpt": 21,
    "modifier": 22,
    "qty": 23,
    "order_date": 25,
    "supervising_md": 26,   # "Supervising MD" column used as supervisor
    "order_md": 27,
    "schedule_staff": 29,
    "location": 30,
    "signed_off_by": 31,
    "primary_insurer": 36,
}

# Outpatient column map (grouped by Order MD)
_OUTPATIENT_COL = {
    "patient": 1,
    "mrn": 3,
    "description": 4,
    "charge_type": 16,
    "cpt": 18,
    "modifier": 19,
    "qty": 22,
    "order_date": 24,
    "supervising_md": 26,   # "Order MD" column used as supervisor for outpatient
    "order_md": 26,
    "schedule_staff": 28,
    "location": 29,
    "signed_off_by": 30,
    "primary_insurer": 35,
}

_GROUP_PATTERNS = {
    "inpatient":  re.compile(r"Supervising MD:", re.IGNORECASE),
    "outpatient": re.compile(r"Order MD:",       re.IGNORECASE),
}

# Legacy alias so existing code that imports COL still works
COL = _INPATIENT_COL


def _find_header_row(df: pd.DataFrame) -> int:
    for i, row in df.iterrows():
        if str(row.iloc[1]).strip() == _HEADER_ROW_MARKER:
            return int(i)
    raise ValueError("Could not locate header row — expected 'Patient' in column B")


def _normalize_cpt(code: str) -> str:
    """Normalize CPT codes that have a modifier appended (e.g. '9921325' → '99213').
    Standard CPT codes are 5 characters; anything longer is treated as CPT + modifier."""
    code = code.strip()
    return code[:5] if len(code) > 5 else code


def parse(path: str, report_type: str = "inpatient") -> pd.DataFrame:
    """
    Parse an iKnowMed All Signed Charges export into a clean DataFrame.

    Parameters
    ----------
    path        : path to the source xlsx
    report_type : "inpatient" (default) or "outpatient"
                  Controls which column positions and group-header pattern are used.
                  In both cases the output DataFrame uses the column name
                  'supervising_md' to hold the physician-supervisor field.
    """
    col = _INPATIENT_COL if report_type == "inpatient" else _OUTPATIENT_COL
    group_pat = _GROUP_PATTERNS.get(report_type, _GROUP_PATTERNS["inpatient"])

    raw = pd.read_excel(path, sheet_name=0, header=None, dtype=str)
    header_idx = _find_header_row(raw)
    data = raw.iloc[header_idx + 1 :].reset_index(drop=True)

    # Drop group-header rows (e.g. "Supervising MD: ..." or "Order MD: ...")
    mask_group = data.iloc[:, col["patient"]].apply(
        lambda v: bool(group_pat.search(str(v)))
    )
    data = data[~mask_group].reset_index(drop=True)

    # Drop rows with no CPT code
    data = data[data.iloc[:, col["cpt"]].notna()].reset_index(drop=True)
    data = data[data.iloc[:, col["cpt"]].str.strip() != "nan"].reset_index(drop=True)

    out = pd.DataFrame(
        {
            "patient": data.iloc[:, col["patient"]].str.strip(),
            "mrn": data.iloc[:, col["mrn"]].str.strip(),
            "description": data.iloc[:, col["description"]].str.strip(),
            "charge_type": data.iloc[:, col["charge_type"]].str.strip(),
            "cpt": data.iloc[:, col["cpt"]].str.strip().apply(_normalize_cpt),
            "modifier": data.iloc[:, col["modifier"]].str.strip(),
            "qty": pd.to_numeric(data.iloc[:, col["qty"]], errors="coerce").fillna(1).astype(int),
            "order_date": data.iloc[:, col["order_date"]].str.strip(),
            "supervising_md": data.iloc[:, col["supervising_md"]].str.strip(),
            "order_md": data.iloc[:, col["order_md"]].str.strip(),
            "location": data.iloc[:, col["location"]].str.strip(),
            "signed_off_by": data.iloc[:, col["signed_off_by"]].str.strip(),
            "primary_insurer": data.iloc[:, col["primary_insurer"]].str.strip(),
        }
    )
    # Normalize NaN strings
    out = out.replace("nan", pd.NA)
    return out


def extract_date_range(path: str) -> str:
    """Extract the date range string from the report header."""
    raw = pd.read_excel(path, sheet_name=0, header=None, dtype=str, nrows=12)
    for _, row in raw.iterrows():
        for cell in row:
            text = str(cell)
            if "Date of Service" in text:
                return text.strip()
    return ""
