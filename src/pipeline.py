from __future__ import annotations

from typing import Any, Dict, List, Tuple
from dataclasses import dataclass
import json

from .filters import read_filter_questions, validate_answers
from .transform import (
    load_input_workbook,
    _get_ws,
    read_start_page_fields,
    find_qualifying_projection_columns,
    extract_transpose_vector,
    build_static_row_values,
)
from .geocode import GeocodeConfig
from .writer import insert_and_write_row

@dataclass
class RunResult:
    rows_written: int
    qualifying_columns: int
    warnings: List[str]

def run_transpose_job(
    input_path: str,
    settings_path: str,
    answers: Dict[int, Any],
) -> RunResult:
    settings = json.loads(open(settings_path, "r", encoding="utf-8").read())

    # Load questions (for validation)
    questions = read_filter_questions(
        output_template_path=settings["output_template_path"],
        filters_sheet_name=settings["filters_sheet_name"],
        question_row=settings["filters"]["question_row"],
        required_row=settings["filters"]["required_row"],
        start_col=settings["filters"]["start_col"],
        end_col=settings["filters"]["end_col"],
        dropdown_start_row=settings["filters"]["dropdown_start_row"],
    )

    ok, missing = validate_answers(questions, answers)
    if not ok:
        raise ValueError("Missing required filters:\n- " + "\n- ".join(missing))

    wb_in = load_input_workbook(input_path)
    ws_start = _get_ws(wb_in, settings["input_sheets"]["start_page"])
    ws_proj = _get_ws(wb_in, settings["input_sheets"]["projections"])

    start_fields = read_start_page_fields(ws_start)

    qual_cols = find_qualifying_projection_columns(
        ws_proj,
        scan_start_col=settings["projections"]["scan_start_col"],
        scan_end_col=settings["projections"]["scan_end_col"],
        last_month_row=settings["projections"]["condition_last_month_row"],
        status_row=settings["projections"]["condition_status_row"],
    )

    geocfg = GeocodeConfig(**settings.get("geocoding", {}))

    static_values, static_warnings = build_static_row_values(start_fields, answers, geocfg)

    rows_written = 0
    warnings = list(static_warnings)

    # Decide processing order:
    # E->U means earliest on top? With insert-at-row-2, the LAST processed ends up on top.
    # We'll process right-to-left so the leftmost qualifying ends up deepest, rightmost ends on top.
    for col_idx in reversed(qual_cols):
        dyn_vec = extract_transpose_vector(
            ws_proj,
            col_idx=col_idx,
            start_row=settings["projections"]["transpose_start_row"],
            end_row=settings["projections"]["transpose_end_row"],
        )

        insert_and_write_row(
            destination_path=settings["destination_workbook_path"],
            sheet_name=settings["destination_sheet_name"],
            insert_row_index=settings["output_layout"]["insert_row_index"],
            static_values_A_to_BC=static_values,
            dynamic_values_BD_to_LZ=dyn_vec,
            static_start_col=settings["output_layout"]["static_start_col"],
            dynamic_start_col=settings["output_layout"]["dynamic_start_col"],
        )
        rows_written += 1

    return RunResult(rows_written=rows_written, qualifying_columns=len(qual_cols), warnings=warnings)
