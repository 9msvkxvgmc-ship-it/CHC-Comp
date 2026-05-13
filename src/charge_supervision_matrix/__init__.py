"""
charge_supervision_matrix
=========================
Process an iKnowMed 'All Signed Charges' Excel export into:
  - Charge Summary: wRVU by supervising MD broken down by CPT code
  - Supervision Matrix: % of each APP's wRVUs attributed to each supervising MD

Quickstart
----------
    from charge_supervision_matrix import run, Config

    run(
        input_path="Q1_charges.xlsx",
        output_path="Q1_report.xlsx",
        config=Config(
            excluded_signers=["Hogan, Jennifer"],
            add_to_app_list=["Tol, Margie"],
        ),
    )
"""

from .config import Config, CHCWM_Q1_2026
from .runner import run

__all__ = ["Config", "CHCWM_Q1_2026", "run"]
