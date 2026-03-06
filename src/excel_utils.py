from __future__ import annotations
from openpyxl.utils import column_index_from_string, get_column_letter

def col_to_idx(col: str) -> int:
    return column_index_from_string(col)

def idx_to_col(idx: int) -> str:
    return get_column_letter(idx)

def normalize_str(x) -> str:
    if x is None:
        return ""
    return str(x).strip()

def is_blank(x) -> bool:
    return normalize_str(x) == ""
