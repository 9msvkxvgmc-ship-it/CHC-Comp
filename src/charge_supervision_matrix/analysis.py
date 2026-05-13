"""
Core calculations: charge summary and supervision matrix.
"""

import re
import pandas as pd

from .config import Config


def _is_app(name: str, config: Config) -> bool:
    if not name or pd.isna(name):
        return False
    name = str(name)
    # Reclassified MDs are never treated as APPs, regardless of credentials
    if name in config.reclassify_as_supervising_md:
        return False
    if name in config.add_to_app_list:
        return True
    for pattern in config.app_credential_patterns:
        if re.search(pattern, name):
            return True
    return False


def _clean_for_summary(df: pd.DataFrame, config: Config) -> pd.DataFrame:
    """
    Light clean for the Charge Summary sheet.

    Only rows for omitted persons are removed entirely.  Charges signed by
    excluded_signers are KEPT and attributed to their supervising MD — this
    captures e.g. admin staff who sign hospital consult charges on behalf of
    a physician.  The signer identity is not surfaced on the summary sheet;
    only supervising_md / cpt / count / wRVU are reported.
    """
    omitted = set(config.omit)

    df = df[~df["signed_off_by"].isin(omitted)].copy()
    df = df[~df["supervising_md"].isin(omitted)].copy()

    if config.location_filter:
        df = df[df["location"].str.contains(config.location_filter, case=False, na=False)].copy()

    return df.reset_index(drop=True)


def _clean(df: pd.DataFrame, config: Config) -> pd.DataFrame:
    """
    Full clean for the Supervision Matrix.

    Drops rows for both excluded_signers and omitted persons — excluded people
    must not appear as APP columns in the matrix.
    """
    all_excluded = set(config.excluded_signers) | set(config.omit)

    # Drop rows where the signer is excluded or omitted
    df = df[~df["signed_off_by"].isin(all_excluded)].copy()

    # Drop rows where the supervising MD is omitted
    df = df[~df["supervising_md"].isin(all_excluded)].copy()

    if config.location_filter:
        df = df[df["location"].str.contains(config.location_filter, case=False, na=False)].copy()

    return df.reset_index(drop=True)


def build_charge_summary(
    df: pd.DataFrame,
    config: Config,
    wrvu_table: dict[str, float],
) -> pd.DataFrame:
    """
    Returns a long-form DataFrame:
      supervising_md | cpt | description | count | wrvu_per_unit | total_wrvu

    All charges for each supervising MD are included regardless of who signed
    them — excluded_signers' charges roll up to the supervising MD rather than
    being dropped.  Only omitted persons are removed entirely.
    """
    df = _clean_for_summary(df, config)

    rows = []
    for (sup_md, cpt), grp in df.groupby(["supervising_md", "cpt"], sort=True):
        count = int(grp["qty"].sum())
        wrvu_unit = wrvu_table.get(str(cpt), 0.0)
        total = round(count * wrvu_unit, 2)
        desc = grp["description"].dropna().mode()
        desc_str = str(desc.iloc[0]) if len(desc) else ""
        rows.append(
            {
                "supervising_md": sup_md,
                "cpt": cpt,
                "description": desc_str,
                "count": count,
                "wrvu_per_unit": wrvu_unit,
                "total_wrvu": total,
            }
        )

    summary = pd.DataFrame(rows)

    # Add per-MD totals
    md_totals = summary.groupby("supervising_md")["total_wrvu"].sum().rename("md_total_wrvu")
    summary = summary.join(md_totals, on="supervising_md")
    return summary


def build_supervision_matrix(
    df: pd.DataFrame,
    config: Config,
    wrvu_table: dict[str, float],
) -> pd.DataFrame:
    """
    Returns a pivot table:
      rows    = supervising MDs (true MDs + reclassified + any APP who only supervises)
      columns = APPs (identified by credentials in signed_off_by)
      values  = fraction of APP's total wRVUs attributed to this supervising MD
    Last column = supervising MD total wRVU (all charges under them)
    Last row    = APP total wRVU signed
    """
    df = _clean(df, config)
    df = df.copy()
    df["wrvu"] = df["cpt"].map(wrvu_table).fillna(0.0) * df["qty"]

    # Identify APP rows (where the signer is an APP)
    df["is_app"] = df["signed_off_by"].apply(lambda n: _is_app(n, config))
    app_df = df[df["is_app"]].copy()

    # Determine which supervising_md values are "true" supervising MDs
    # (not APPs who also sign their own charges under themselves)
    # Rule: a person is a true supervising MD if they are NOT an APP-signer,
    # OR if they are in reclassify_as_supervising_md.
    app_signers = set(df.loc[df["is_app"], "signed_off_by"].dropna().unique())
    reclassified = set(config.reclassify_as_supervising_md)

    def _is_true_sup_md(name: str) -> bool:
        if not name or pd.isna(name):
            return False
        name = str(name)
        # Reclassified MDs are always true supervising MDs
        if name in reclassified:
            return True
        # Anyone who is an active APP signer belongs in columns, not rows
        # (unless they're reclassified, handled above)
        if name in app_signers:
            return False
        return True

    # Pivot: wRVUs by (supervising_md, app_signer) — only APP charges
    pivot = (
        app_df.groupby(["supervising_md", "signed_off_by"])["wrvu"]
        .sum()
        .unstack(fill_value=0.0)
    )

    # Filter pivot rows to true supervising MDs only
    true_sup_rows = [md for md in pivot.index if _is_true_sup_md(md)]
    pivot = pivot.loc[true_sup_rows]

    # Sort alphabetically
    pivot = pivot.reindex(sorted(pivot.columns), axis=1)
    pivot = pivot.reindex(sorted(pivot.index))

    # Convert to proportions: each APP column / APP's total across ALL rows
    # (denominator = total wRVUs that APP signed across all supervising MDs,
    #  including any rows we filtered out — so we recompute from the full app_df)
    app_totals_full = (
        app_df.groupby("signed_off_by")["wrvu"]
        .sum()
        .reindex(pivot.columns, fill_value=0.0)
    )
    matrix = pivot.div(app_totals_full, axis=1).round(6)
    matrix = matrix.replace([float("inf"), float("-inf")], 0.0)

    # Supervising MD total wRVU (ALL charges under them, not just APP charges)
    md_total_wrvu = (
        df.groupby("supervising_md")["wrvu"]
        .sum()
        .reindex(matrix.index, fill_value=0.0)
        .round(2)
    )
    matrix["Total wRVU (Sup MD)"] = md_total_wrvu

    # APP totals row
    app_totals_row = app_totals_full.round(2).to_dict()
    app_totals_row["Total wRVU (Sup MD)"] = None
    matrix.loc["APP Total wRVU Signed"] = app_totals_row

    return matrix
