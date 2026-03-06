from __future__ import annotations

from typing import Any, List, Tuple
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from .excel_utils import col_to_idx, idx_to_col

def insert_and_write_row(
    destination_path: str,
    sheet_name: str,
    insert_row_index: int,
    static_values_A_to_BC: List[Any],
    dynamic_values_BD_to_LZ: List[Any],
    static_start_col: str = "A",
    dynamic_start_col: str = "BD",
) -> None:
    wb = load_workbook(destination_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Destination workbook missing sheet: {sheet_name}")
    ws = wb[sheet_name]

    # Insert row at row 2 (Option A)
    ws.insert_rows(insert_row_index, amount=1)

    # Write static A:BC by position
    start_idx = col_to_idx(static_start_col)
    for offset, val in enumerate(static_values_A_to_BC):
        ws.cell(row=insert_row_index, column=start_idx + offset).value = val

    # Write dynamic BD:LZ by position
    dyn_start_idx = col_to_idx(dynamic_start_col)
    for offset, val in enumerate(dynamic_values_BD_to_LZ):
        ws.cell(row=insert_row_index, column=dyn_start_idx + offset).value = val

    wb.save(destination_path)
