"""
High-level run() function — the main entry point for programmatic use.
"""

import os
from .config import Config
from . import parser, wrvu, analysis, output


def run(
    input_path: str,
    output_path: str | None = None,
    config: Config | None = None,
    report_type: str = "inpatient",
) -> str:
    """
    Process a raw 'All Signed Charges' Excel export and write an output workbook.

    Parameters
    ----------
    input_path  : path to the source xlsx (All Signed Charges report)
    output_path : where to write the output xlsx; defaults to <stem>_report.xlsx
                  next to the input file
    config      : Config instance; defaults to Config() with no exclusions
    report_type : "inpatient" (default) or "outpatient"

    Returns
    -------
    Absolute path to the written output file.
    """
    if config is None:
        config = Config()

    if output_path is None:
        stem = os.path.splitext(os.path.abspath(input_path))[0]
        output_path = stem + "_report.xlsx"

    wrvu_table = wrvu.load(config.wrvu_file)
    df = parser.parse(input_path, report_type=report_type)
    date_range = parser.extract_date_range(input_path)
    source_name = os.path.basename(input_path)

    summary = analysis.build_charge_summary(df, config, wrvu_table)
    matrix = analysis.build_supervision_matrix(df, config, wrvu_table)

    output.write_workbook(
        output_path=output_path,
        summary=summary,
        matrix=matrix,
        config=config,
        date_range=date_range,
        source_name=source_name,
    )

    return os.path.abspath(output_path)
