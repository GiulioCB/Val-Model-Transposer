from __future__ import annotations

from dataclasses import replace
from typing import List, Dict, Any, Tuple
from .excel_utils import normalize_str
from .types import FilterQuestion

HARDCODED_FILTER_QUESTIONS: List[FilterQuestion] = [
    FilterQuestion(idx=1, title="PropertyType", required=False, kind="dropdown", options=[]),
    FilterQuestion(idx=2, title="Location", required=False, kind="dropdown", options=[]),
    FilterQuestion(idx=3, title="Food Bev Operator", required=False, kind="dropdown", options=[]),
    FilterQuestion(idx=4, title="Operator", required=False, kind="dropdown", options=[]),
    FilterQuestion(idx=5, title="Chain/ChainID", required=False, kind="dropdown", options=[]),
    FilterQuestion(idx=6, title="Management Company", required=False, kind="dropdown", options=[]),
    FilterQuestion(idx=7, title="Owner Company", required=False, kind="dropdown", options=[]),
    FilterQuestion(idx=8, title="YearOpened", required=False, kind="text", options=None),
    FilterQuestion(idx=9, title="MeetingSpace (SQM)", required=False, kind="text", options=None),
    FilterQuestion(idx=10, title="Ski", required=False, kind="yesno", options=["Yes", "No"]),
    FilterQuestion(idx=11, title="Spa", required=False, kind="yesno", options=["Yes", "No"]),
    FilterQuestion(idx=12, title="HealthClub", required=False, kind="yesno", options=["Yes", "No"]),
    FilterQuestion(idx=13, title="Golf", required=False, kind="yesno", options=["Yes", "No"]),
    FilterQuestion(idx=14, title="Boutique", required=False, kind="yesno", options=["Yes", "No"]),
    FilterQuestion(idx=15, title="FoodOutlets", required=False, kind="text", options=None),
    FilterQuestion(idx=16, title="BeverageOutlets", required=False, kind="text", options=None),
]

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
    return [replace(question) for question in HARDCODED_FILTER_QUESTIONS]

def validate_answers(questions: List[FilterQuestion], answers: Dict[int, Any]) -> Tuple[bool, List[str]]:
    missing: List[str] = []
    for q in questions:
        if not q.required:
            continue
        v = answers.get(q.idx, None)
        if q.kind in ("dropdown", "text"):
            if v is None or normalize_str(v) == "":
                missing.append(f"Filter {q.idx}: {q.title}")
        elif q.kind == "yesno":
            if v not in ("Yes", "No"):
                missing.append(f"Filter {q.idx}: {q.title}")
        else:
            if v is None or normalize_str(v) == "":
                missing.append(f"Filter {q.idx}: {q.title}")
    return (len(missing) == 0, missing)
