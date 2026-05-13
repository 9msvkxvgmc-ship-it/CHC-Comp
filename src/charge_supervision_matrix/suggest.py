"""
Analyze a raw charge file and suggest --exclude / --reclassify-as-md / --add-app flags.
"""

import re
import pandas as pd

from . import parser

_MD_PATTERN = re.compile(r"\b(MD|DO|MD PhD|PhD MD)\b")
_APP_PATTERN = re.compile(r"\b(PA-C|PA|NP|AGNP-C|CRNP|FNP|ANP|DNP|FNP-BC|NP-C)\b")
_OTHER_PATTERN = re.compile(r"\b(PhD|RN|RN-BC|LCSW|LPN|CMA|MT|RT|RPh|PharmD)\b")


def _classify(name: str) -> str:
    if _MD_PATTERN.search(name):
        return "md"
    if _APP_PATTERN.search(name):
        return "app"
    if _OTHER_PATTERN.search(name):
        return "other"
    return "unknown"


def analyze(path: str) -> dict:
    df = parser.parse(path)

    signers = sorted(df["signed_off_by"].dropna().unique())
    charge_counts = df["signed_off_by"].value_counts().to_dict()

    categorized = {name: _classify(name) for name in signers}

    md_signers    = [n for n, c in categorized.items() if c == "md"]
    app_signers   = [n for n, c in categorized.items() if c == "app"]
    other_signers = [n for n, c in categorized.items() if c == "other"]
    unknown_signers = [n for n, c in categorized.items() if c == "unknown"]

    # For each unrecognized/other signer, find which supervising MDs they are
    # the sole (or primary) signer for.  If excluding them would make a supervising
    # MD vanish from the matrix entirely, warn and recommend --add-app instead.
    sole_signer_for: dict[str, list[str]] = {}
    for signer in other_signers + unknown_signers:
        signer_rows = df[df["signed_off_by"] == signer]
        sup_mds_for_signer = signer_rows["supervising_md"].dropna().unique().tolist()
        orphaned = []
        for sup_md in sup_mds_for_signer:
            sup_md_rows = df[df["supervising_md"] == sup_md]
            other_signers_for_md = sup_md_rows[
                sup_md_rows["signed_off_by"] != signer
            ]["signed_off_by"].dropna().unique()
            # Only count non-MD other signers as keeping this MD "alive"
            non_md_others = [s for s in other_signers_for_md if _classify(s) != "md"]
            if len(non_md_others) == 0:
                orphaned.append(sup_md)
        if orphaned:
            sole_signer_for[signer] = sorted(orphaned)

    return {
        "md_signers": md_signers,
        "app_signers": app_signers,
        "other_signers": other_signers,
        "unknown_signers": unknown_signers,
        "sole_signer_for": sole_signer_for,
        "charge_counts": charge_counts,
        "date_range": parser.extract_date_range(path),
        "total_charges": len(df),
    }


def print_report(path: str):
    r = analyze(path)

    print(f"\nFile: {path}")
    print(f"Date range: {r['date_range']}")
    print(f"Total charges parsed: {r['total_charges']}\n")

    print("=" * 60)
    print("MDs signing charges  →  consider --reclassify-as-md")
    print("=" * 60)
    if r["md_signers"]:
        for name in r["md_signers"]:
            print(f"  {name}  ({r['charge_counts'].get(name, 0)} charges)")
    else:
        print("  (none found)")

    print()
    print("=" * 60)
    print("APPs signing charges  →  will appear as matrix columns")
    print("=" * 60)
    for name in r["app_signers"]:
        print(f"  {name}  ({r['charge_counts'].get(name, 0)} charges)")

    safe_exclude  = []
    add_app_recommended = []
    for name in r["other_signers"] + r["unknown_signers"]:
        if name in r["sole_signer_for"]:
            add_app_recommended.append(name)
        else:
            safe_exclude.append(name)

    if add_app_recommended:
        print()
        print("=" * 60)
        print("⚠  No-credential signers who are the SOLE signer for certain")
        print("   supervising MDs  →  use --add-app (NOT --exclude)")
        print("   Excluding these will remove those MDs from the matrix entirely.")
        print("=" * 60)
        for name in add_app_recommended:
            n = r["charge_counts"].get(name, 0)
            mds = r["sole_signer_for"][name]
            print(f"  {name}  ({n} charges)")
            print(f"    Sole/primary signer for: {', '.join(mds)}")

    if safe_exclude:
        print()
        print("=" * 60)
        print("Other / unrecognized credentials  →  consider --exclude")
        print("=" * 60)
        for name in safe_exclude:
            print(f"  {name}  ({r['charge_counts'].get(name, 0)} charges)")

    print()
    print("=" * 60)
    print("Suggested command (edit as needed)")
    print("=" * 60)
    print(f'charge-supervision-matrix "{path}" \\')
    for name in r["md_signers"]:
        print(f'  --reclassify-as-md "{name}" \\')
    for name in add_app_recommended:
        print(f'  --add-app "{name}" \\')
    for name in safe_exclude:
        print(f'  --exclude "{name}" \\')
    print('  -o "output.xlsx"')
    print()
