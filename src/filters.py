from __future__ import annotations

import json
from dataclasses import replace
from pathlib import Path
from typing import Any, Dict, List, Tuple

from .excel_utils import normalize_str
from .types import FilterQuestion

FILTER_OPTIONS_PATH = Path(__file__).resolve().parent.parent / "configs" / "filter_options.json"
FILTER_5_OPTIONS_PATH = Path(__file__).resolve().parent.parent / "configs" / "filter_5_options.json"
FILTER_6_OPTIONS_PATH = Path(__file__).resolve().parent.parent / "configs" / "filter_6_options.json"
FILTER_7_OPTIONS_PATH = Path(__file__).resolve().parent.parent / "configs" / "filter_7_options.json"

HARDCODED_FILTER_QUESTIONS: List[FilterQuestion] = [
    FilterQuestion(
        idx=1,
        title="PropertyType",
        required=True,
        kind="fixed_dropdown",
        options=[
            "All-Inclusive Hotel",
            "Aparthotel",
            "Full-Service Hotel",
            "Convention Hotel",
            "Hostel",
            "Limited-Service Hotel",
            "Resort",
        ],
    ),
    FilterQuestion(
        idx=2,
        title="Location",
        required=True,
        kind="fixed_dropdown",
        options=[
            "Airport",
            "City Centre",
            "Highway",
            "Peripheral",
            "Resort",
            "Rural/Countryside",
            "Seaside",
            "Small Metro/Town",
            "Suburban",
        ],
    ),
    FilterQuestion(
        idx=3,
        title="Food Bev Operator",
        required=False,
        kind="fixed_dropdown",
        options=[
            "Combination",
            "None or Offsite",
            "Onsite, Hotel Operated",
            "Onsite, Leased Out",
            "Unassigned",
        ],
    ),
    FilterQuestion(
        idx=4,
        title="Operator",
        required=True,
        kind="fixed_dropdown",
        options=[
            "Franchise/Chain Operator",
            "Management Company",
            "Operating Lease",
            "Owner Operated",
        ],
    ),
    FilterQuestion(
        idx=5,
        title="Chain/ChainID",
        required=False,
        kind="dropdown",
        options=[],
    ),
    FilterQuestion(
        idx=6,
        title="Management Company",
        required=False,
        kind="dropdown",
        options=[],
    ),
    FilterQuestion(
        idx=7,
        title="Owner Company",
        required=False,
        kind="dropdown",
        options=[],
    ),
    FilterQuestion(idx=8, title="YearOpened", required=False, kind="date", options=None),
    FilterQuestion(idx=9, title="MeetingSpace (SQM)", required=False, kind="number", options=None),
    FilterQuestion(idx=10, title="Ski", required=True, kind="yesno", options=["Yes", "No"]),
    FilterQuestion(idx=11, title="Spa", required=True, kind="yesno", options=["Yes", "No"]),
    FilterQuestion(idx=12, title="HealthClub", required=True, kind="yesno", options=["Yes", "No"]),
    FilterQuestion(idx=13, title="Golf", required=True, kind="yesno", options=["Yes", "No"]),
    FilterQuestion(idx=14, title="Boutique", required=True, kind="yesno", options=["Yes", "No"]),
    FilterQuestion(idx=15, title="FoodOutlets", required=True, kind="number", options=None),
    FilterQuestion(idx=16, title="BeverageOutlets", required=True, kind="number", options=None),
]


def _dedupe_options(options: List[str]) -> List[str]:
    seen: set[str] = set()
    cleaned: List[str] = []

    for option in options:
        value = normalize_str(option)
        if not value:
            continue

        key = value.casefold()
        if key in seen:
            continue

        seen.add(key)
        cleaned.append(value)

    return cleaned


def _load_saved_dropdown_options() -> Dict[str, List[str]]:
    if not FILTER_OPTIONS_PATH.exists():
        return {}

    with FILTER_OPTIONS_PATH.open("r", encoding="utf-8") as fh:
        data = json.load(fh)

    if not isinstance(data, dict):
        return {}

    saved_options: Dict[str, List[str]] = {}
    for key, values in data.items():
        if isinstance(values, list):
            saved_options[str(key)] = _dedupe_options([str(value) for value in values])

    return saved_options


def _load_external_filter_options(path: Path, expected_title: str) -> List[str]:
    if not path.exists():
        return []

    with path.open("r", encoding="utf-8") as fh:
        data = json.load(fh)

    if not isinstance(data, dict):
        return []

    if normalize_str(data.get("title")) != expected_title:
        return []

    options = data.get("options", [])
    if not isinstance(options, list):
        return []

    return _dedupe_options([str(option) for option in options])


def normalize_answers_for_processing(answers: Dict[int, Any]) -> Dict[int, Any]:
    normalized_answers = dict(answers)

    year_opened = normalize_str(normalized_answers.get(8))
    if year_opened:
        normalized_answers[8] = f"{year_opened}-01-01"

    return normalized_answers


def _save_dropdown_options(options_by_title: Dict[str, List[str]]) -> None:
    FILTER_OPTIONS_PATH.parent.mkdir(parents=True, exist_ok=True)

    serializable = {
        title: _dedupe_options(values)
        for title, values in sorted(options_by_title.items())
    }

    with FILTER_OPTIONS_PATH.open("w", encoding="utf-8") as fh:
        json.dump(serializable, fh, indent=2)


def read_filter_questions(
    output_template_path: str | None = None,
    filters_sheet_name: str | None = None,
    question_row: int | None = None,
    required_row: int | None = None,
    start_col: str | None = None,
    end_col: str | None = None,
    dropdown_start_row: int | None = None,
) -> List[FilterQuestion]:
    # Kept signature-compatible with the previous workbook-based implementation so
    # existing callers do not need to change when switching to code-defined filters.
    saved_options = _load_saved_dropdown_options()
    questions: List[FilterQuestion] = []

    for question in HARDCODED_FILTER_QUESTIONS:
        cloned = replace(question)
        if cloned.idx == 5:
            cloned.options = _load_external_filter_options(FILTER_5_OPTIONS_PATH, cloned.title)
        elif cloned.idx == 6:
            cloned.options = _load_external_filter_options(FILTER_6_OPTIONS_PATH, cloned.title)
        elif cloned.idx == 7:
            cloned.options = _load_external_filter_options(FILTER_7_OPTIONS_PATH, cloned.title)
        if cloned.kind in ("dropdown", "fixed_dropdown"):
            merged_options = list(cloned.options or [])
            merged_options.extend(saved_options.get(cloned.title, []))
            cloned.options = _dedupe_options(merged_options)
        questions.append(cloned)

    return questions


def add_dropdown_option(question_title: str, new_option: str) -> str | None:
    normalized_title = normalize_str(question_title)
    normalized_option = normalize_str(new_option)

    if not normalized_title or not normalized_option:
        return None

    question_lookup = {
        normalize_str(question.title).casefold(): question
        for question in HARDCODED_FILTER_QUESTIONS
        if question.kind == "dropdown"
    }
    question = question_lookup.get(normalized_title.casefold())
    if question is None:
        raise ValueError(f"Dropdown filter '{question_title}' was not found.")

    saved_options = _load_saved_dropdown_options()
    merged_options = list(question.options or [])
    merged_options.extend(saved_options.get(question.title, []))

    existing_before_save = {
        option.casefold()
        for option in (question.options or []) + saved_options.get(question.title, [])
    }
    if normalized_option.casefold() in existing_before_save:
        return None

    merged_options.append(normalized_option)
    deduped_options = _dedupe_options(merged_options)
    saved_options[question.title] = deduped_options
    _save_dropdown_options(saved_options)
    return normalized_option


def validate_answers(questions: List[FilterQuestion], answers: Dict[int, Any]) -> Tuple[bool, List[str]]:
    problems: List[str] = []
    for q in questions:
        v = answers.get(q.idx, None)
        value_str = normalize_str(v)

        if q.required:
            if q.kind in ("dropdown", "fixed_dropdown", "text", "number", "date"):
                if value_str == "":
                    problems.append(f"Filter {q.idx}: {q.title}")
                    continue
            elif q.kind == "yesno":
                if v not in ("Yes", "No"):
                    problems.append(f"Filter {q.idx}: {q.title}")
                    continue
            else:
                if value_str == "":
                    problems.append(f"Filter {q.idx}: {q.title}")
                    continue

        if q.kind == "number" and value_str != "":
            try:
                float(value_str.replace(",", ""))
            except ValueError:
                problems.append(f"Filter {q.idx}: {q.title} must be a number")
        elif q.kind == "date" and value_str != "":
            if len(value_str) != 4 or not value_str.isdigit():
                problems.append(f"Filter {q.idx}: {q.title} must be a year in YYYY format")

    return (len(problems) == 0, problems)
