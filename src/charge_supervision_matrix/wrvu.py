"""
wRVU lookup from CMS Physician Fee Schedule.

Built-in table covers common inpatient/E&M codes (2026 MPFS Final Rule, RVU26A).
Pass a CMS RVU CSV file path to override or extend.
"""

import csv

_BUILTIN: dict[str, float] = {
    # Initial hospital care
    "99221": 2.65,
    "99222": 2.61,
    "99223": 3.86,
    # Subsequent hospital care
    "99231": 0.76,
    "99232": 1.39,
    "99233": 2.00,
    # Hospital discharge
    "99238": 1.28,
    "99239": 1.90,
    # Critical care
    "99291": 4.50,
    "99292": 2.25,
    # Observation
    "99217": 1.50,
    "99218": 1.92,
    "99219": 2.64,
    "99220": 3.57,
    "99234": 2.33,
    "99235": 3.17,
    "99236": 4.20,
    # ED
    "99281": 0.45,
    "99282": 0.88,
    "99283": 1.34,
    "99284": 2.56,
    "99285": 3.80,
    # Outpatient/Office new
    "99202": 0.93,
    "99203": 1.60,
    "99204": 2.60,
    "99205": 3.50,
    # Outpatient/Office established
    "99211": 0.18,
    "99212": 0.70,
    "99213": 1.30,
    "99214": 1.92,
    "99215": 2.80,
    # Prolonged services
    "99417": 1.62,
    "99418": 0.81,
    "994X0": 0.78,
    # Chemo admin / infusion (wRVU = 0 for technical; included for completeness)
    "96542": 0.0,
    # Bone marrow biopsy/aspiration
    "38220": 1.17,
    "38221": 1.25,
    "38222": 1.40,
    # Psychiatric/therapy
    "90791": 2.79,
    "90792": 2.79,
    "90832": 0.87,
    "90833": 0.87,
    "90834": 1.35,
    "90836": 1.35,
    "90837": 1.94,
    "90838": 1.94,
    # Consultation
    "99241": 0.97,
    "99242": 1.62,
    "99243": 2.27,
    "99244": 3.36,
    "99245": 4.20,
    "99251": 1.01,
    "99252": 1.54,
    "99253": 2.25,
    "99254": 3.18,
    "99255": 4.20,
}


def load(path: str | None = None) -> dict[str, float]:
    table = dict(_BUILTIN)
    if path:
        with open(path, newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                cpt = row.get("HCPCS") or row.get("CPT") or row.get("hcpcs") or ""
                wrvu = row.get("WORK RVU") or row.get("work_rvu") or row.get("WORK_RVU") or "0"
                cpt = cpt.strip()
                try:
                    val = float(wrvu.strip())
                except ValueError:
                    continue
                if cpt:
                    table[cpt] = val
    return table
