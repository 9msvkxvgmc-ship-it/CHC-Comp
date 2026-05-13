"""
Parse a raw 'All Signed Charges' Excel export into a clean DataFrame.

The export has a multi-row header section before the actual column headers.
Column positions (0-indexed):
  1  = Patient name
  3  = MRN
  4  = Description (charge description)
  18 = Charge Type
  21 = CPT code
  22 = Modifier
  23 = Qty
  25 = Order Date
  26 = Supervising MD
  27 = Order MD
  29 = Schedule Provider/Staff
  30 = Location
  31 = Signed-Off By
  36 = Primary Insurer
"""

import re
import pandas as pd

_HEADER_ROW_MARKER = "Patient"  # value in col 1 that identifies the real column header row
_GROUP_HEADER_PATTERN = re.compile(r"Supervising MD:", re.IGNORECASE)

COL = {
    "patient": 1,
    "mrn": 3,
    "description": 4,
    "charge_type": 18,
    "cpt": 21,
    "modifier": 22,
    "qty": 23,
    "order_date": 25,
    "supervising_md": 26,
    "order_md": 27,
    "schedule_staff": 29,
    "location": 30,
    "signed_off_by": 31,
    "primary_insurer": 36,
}


def _find_header_row(df: pd.DataFrame) -> int:
    for i, row in df.iterrows():
        if str(row.iloc[1]).strip() == _HEADER_ROW_MARKER:
            return int(i)
    raise ValueError("Could not locate header row — expected 'Patient' in column B")


def parse(path: str) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=0, header=None, dtype=str)
    header_idx = _find_header_row(raw)
    data = raw.iloc[header_idx + 1 :].reset_index(drop=True)

    # Drop group-header rows (e.g. "Supervising MD: Alguire MD, Kathryn")
    mask_group = data.iloc[:, COL["patient"]].apply(
        lambda v: bool(_GROUP_HEADER_PATTERN.search(str(v)))
    )
    data = data[~mask_group].reset_index(drop=True)

    # Drop rows with no CPT code
    data = data[data.iloc[:, COL["cpt"]].notna()].reset_index(drop=True)
    data = data[data.iloc[:, COL["cpt"]].str.strip() != "nan"].reset_index(drop=True)

    out = pd.DataFrame(
        {
            "patient": data.iloc[:, COL["patient"]].str.strip(),
            "mrn": data.iloc[:, COL["mrn"]].str.strip(),
            "description": data.iloc[:, COL["description"]].str.strip(),
            "charge_type": data.iloc[:, COL["charge_type"]].str.strip(),
            "cpt": data.iloc[:, COL["cpt"]].str.strip(),
            "modifier": data.iloc[:, COL["modifier"]].str.strip(),
            "qty": pd.to_numeric(data.iloc[:, COL["qty"]], errors="coerce").fillna(1).astype(int),
            "order_date": data.iloc[:, COL["order_date"]].str.strip(),
            "supervising_md": data.iloc[:, COL["supervising_md"]].str.strip(),
            "order_md": data.iloc[:, COL["order_md"]].str.strip(),
            "location": data.iloc[:, COL["location"]].str.strip(),
            "signed_off_by": data.iloc[:, COL["signed_off_by"]].str.strip(),
            "primary_insurer": data.iloc[:, COL["primary_insurer"]].str.strip(),
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
