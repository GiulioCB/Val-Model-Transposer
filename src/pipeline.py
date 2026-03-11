from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Any, Dict, List, Tuple

from .excel_utils import col_to_idx
from .filters import normalize_answers_for_processing, read_filter_questions, validate_answers
from .geocode import GeocodeConfig
from .transform import (
    _get_ws,
    build_static_row_values,
    extract_transpose_vector,
    find_qualifying_projection_columns,
    load_input_workbook,
    read_projection_row_headers,
    read_start_page_fields,
)
from .writer import (
    insert_and_write_row_in_workbook,
    open_destination_workbook,
    save_destination_workbook,
)


@dataclass
class RunResult:
    rows_written: int
    qualifying_columns: int
    warnings: List[str]


@dataclass
class PreviewResult:
    headers: List[str]
    rows: List[List[Any]]
    qualifying_columns: int
    warnings: List[str]


STATIC_PREVIEW_HEADERS = [
    "Property Name",
    "Street Number",
    "Street Address",
    "Street Type",
    "Street Prefix",
    "Street Suffix",
    "POBox",
    "Country",
    "Country - Region",
    "Region",
    "County",
    "City/Town",
    "District",
    "Zip/Post Code",
    "PropertyType",
    "Owner Type",
    "Location",
    "Food Bev Operator",
    "Operator",
    "Rooms",
    "Metro Area",
    "Market Area",
    "Submarket Area",
    "Chain/ChainID",
    "Classification",
    "Management Company",
    "Owner Company",
    "YearOpened",
    "YearClosed",
    "Year Recent Renovation",
    "OwnBuildings",
    "OwnLand",
    "MeetingSpace (SQM)",
    "MeetingRooms",
    "MeetingMaxCapacity (theatre style only)",
    "Casino",
    "Convention",
    "Conference",
    "Ski",
    "Spa",
    "HealthClub",
    "Golf",
    "Boutique",
    "AllSuite",
    "Suites",
    "Floors",
    "ParkingSpaces",
    "FoodOutlets",
    "BeverageOutlets",
    "Financial Data provider (Source of Information)",
    "Notes",
    "pcd2",
    "lat",
    "long",
    "Currency",
]


def _load_settings(settings_path: str) -> Dict[str, Any]:
    return json.loads(open(settings_path, "r", encoding="utf-8").read())


def _build_preview_headers(
    settings: Dict[str, Any],
    static_value_count: int,
    dynamic_headers: List[Any],
) -> List[str]:
    static_end_idx = col_to_idx(settings["output_layout"]["static_end_col"])
    static_padding_count = max(0, static_end_idx - static_value_count)
    padded_static_headers = STATIC_PREVIEW_HEADERS + ([""] * static_padding_count)
    normalized_dynamic_headers = [
        header if header not in (None, "") else f"Projection Row {idx}"
        for idx, header in enumerate(
            dynamic_headers,
            start=settings["projections"]["transpose_start_row"],
        )
    ]
    return _make_unique_headers(padded_static_headers + normalized_dynamic_headers)


def _make_unique_headers(headers: List[Any]) -> List[str]:
    seen: Dict[str, int] = {}
    unique_headers: List[str] = []

    for idx, header in enumerate(headers, start=1):
        base_header = str(header) if header not in (None, "") else f"Column {idx}"
        count = seen.get(base_header, 0)
        if count == 0:
            unique_headers.append(base_header)
        else:
            unique_headers.append(f"{base_header} ({count + 1})")
        seen[base_header] = count + 1

    return unique_headers


def _prepare_transpose_data(
    input_path: str,
    settings: Dict[str, Any],
    answers: Dict[int, Any],
) -> Tuple[Dict[str, Any], List[List[Any]], List[Any], List[Any], List[str]]:
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

    normalized_answers = normalize_answers_for_processing(answers)

    wb_in = load_input_workbook(input_path)
    ws_start = _get_ws(wb_in, settings["input_sheets"]["start_page"])
    ws_proj = _get_ws(wb_in, settings["input_sheets"]["projections"])

    start_fields = read_start_page_fields(ws_start)
    qualifying_cols = find_qualifying_projection_columns(
        ws_proj,
        scan_start_col=settings["projections"]["scan_start_col"],
        scan_end_col=settings["projections"]["scan_end_col"],
        last_month_row=settings["projections"]["condition_last_month_row"],
        status_row=settings["projections"]["condition_status_row"],
    )

    geocfg = GeocodeConfig(**settings.get("geocoding", {}))
    static_values, warnings = build_static_row_values(start_fields, normalized_answers, geocfg)
    dynamic_rows = [
        extract_transpose_vector(
            ws_proj,
            col_idx=col_idx,
            start_row=settings["projections"]["transpose_start_row"],
            end_row=settings["projections"]["transpose_end_row"],
        )
        for col_idx in qualifying_cols
    ]
    dynamic_headers = read_projection_row_headers(
        ws_proj,
        start_row=settings["projections"]["transpose_start_row"],
        end_row=settings["projections"]["transpose_end_row"],
    )

    return settings, dynamic_rows, static_values, dynamic_headers, list(warnings)


def _open_destination_session(settings: Dict[str, Any]):
    output_storage = settings.get("output_storage", {})
    storage_type = output_storage.get("type", "local")
    return open_destination_workbook(
        destination_path=settings.get("destination_workbook_path") if storage_type == "local" else None,
        github_repo=output_storage.get("github_repo") if storage_type == "github" else None,
        github_excel_path=output_storage.get("github_excel_path") if storage_type == "github" else None,
        github_branch=output_storage.get("github_branch") if storage_type == "github" else None,
    )


def build_preview_result(
    input_path: str,
    settings_path: str,
    answers: Dict[int, Any],
) -> PreviewResult:
    settings = _load_settings(settings_path)
    settings, dynamic_rows, static_values, dynamic_headers, warnings = _prepare_transpose_data(
        input_path, settings, answers
    )

    static_end_idx = col_to_idx(settings["output_layout"]["static_end_col"])
    static_padding = [None] * max(0, static_end_idx - len(static_values))
    rows = [static_values + static_padding + dynamic_values for dynamic_values in dynamic_rows]
    headers = _build_preview_headers(settings, len(static_values), dynamic_headers)

    return PreviewResult(
        headers=headers,
        rows=rows,
        qualifying_columns=len(dynamic_rows),
        warnings=warnings,
    )


def run_transpose_job(
    input_path: str,
    settings_path: str,
    answers: Dict[int, Any],
) -> RunResult:
    preview = build_preview_result(input_path, settings_path, answers)
    settings = _load_settings(settings_path)
    destination_session = _open_destination_session(settings)

    rows_written = 0
    static_end_idx = col_to_idx(settings["output_layout"]["static_end_col"])

    for row_values in reversed(preview.rows):
        insert_and_write_row_in_workbook(
            workbook=destination_session.workbook,
            sheet_name=settings["destination_sheet_name"],
            insert_row_index=settings["output_layout"]["insert_row_index"],
            static_values_A_to_BC=row_values[:static_end_idx],
            dynamic_values_BD_to_LZ=row_values[static_end_idx:],
            static_start_col=settings["output_layout"]["static_start_col"],
            dynamic_start_col=settings["output_layout"]["dynamic_start_col"],
        )
        rows_written += 1

    save_destination_workbook(destination_session)
    return RunResult(
        rows_written=rows_written,
        qualifying_columns=preview.qualifying_columns,
        warnings=preview.warnings,
    )
