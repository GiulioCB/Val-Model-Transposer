from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parents[1]
OUTPUT_PATH = ROOT / "docs" / "output_mapping_template.xlsx"

STATIC_HEADERS = [
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

STATIC_SOURCES = [
    "Start Page!C4",
    "Blank",
    "Start Page!C5",
    "Blank",
    "Blank",
    "Blank",
    "Blank",
    "Blank",
    "Start Page!C8",
    "Blank",
    "Blank",
    "Start Page!C6",
    "Blank",
    "Start Page!C7",
    "Filter 1",
    "Blank",
    "Filter 2",
    "Filter 3",
    "Filter 4",
    "Start Page!C11",
    "Blank",
    "Blank",
    "Blank",
    "Filter 5",
    "Start Page!F10",
    "Filter 6",
    "Filter 7",
    "Filter 8",
    "Blank",
    "Blank",
    "Blank",
    "Blank",
    "Filter 9",
    "Start Page!F15",
    "Start Page!F16",
    "Blank",
    "Blank",
    "Blank",
    "Filter 10",
    "Filter 11",
    "Filter 12",
    "Filter 13",
    "Filter 14",
    "Blank",
    "Blank",
    "Blank",
    "Blank",
    "Filter 15",
    "Filter 16",
    "Blank",
    "Blank",
    "Start Page!C7",
    "Automation 1",
    "Automation 2",
    "Start Page!C21",
]

STATIC_RULES = [
    "Copy property name from the input workbook.",
    "Leave blank.",
    "Copy street address from the input workbook.",
    "Leave blank.",
    "Leave blank.",
    "Leave blank.",
    "Leave blank.",
    "Leave blank.",
    "Copy country/region from the input workbook.",
    "Leave blank.",
    "Leave blank.",
    "Copy city/town from the input workbook.",
    "Leave blank.",
    "Copy zip or post code from the input workbook.",
    "Use the PropertyType answer.",
    "Leave blank.",
    "Use the Location answer.",
    "Use the Food Bev Operator answer.",
    "Use the Operator answer.",
    "Copy rooms from the input workbook.",
    "Leave blank.",
    "Leave blank.",
    "Leave blank.",
    "Use the Chain/ChainID answer.",
    "Copy classification from the input workbook.",
    "Use the Management Company answer.",
    "Use the Owner Company answer.",
    "Use the YearOpened answer.",
    "Leave blank.",
    "Leave blank.",
    "Leave blank.",
    "Leave blank.",
    "Use the MeetingSpace (SQM) answer.",
    "Copy meeting rooms from the input workbook.",
    "Copy meeting max capacity from the input workbook.",
    "Leave blank.",
    "Leave blank.",
    "Leave blank.",
    "Use the Ski answer.",
    "Use the Spa answer.",
    "Use the HealthClub answer.",
    "Use the Golf answer.",
    "Use the Boutique answer.",
    "Leave blank.",
    "Leave blank.",
    "Leave blank.",
    "Leave blank.",
    "Use the FoodOutlets answer.",
    "Use the BeverageOutlets answer.",
    "Leave blank.",
    "Leave blank.",
    "Copy the same value used for Zip/Post Code.",
    "Populate latitude using Automation 1.",
    "Populate longitude using Automation 2.",
    "Copy currency from the input workbook.",
]


def build_output_columns():
    output_columns = []

    for index in range(55):
        excel_col = get_column_letter(index + 1)
        if index < len(STATIC_HEADERS):
            is_blank_source = STATIC_SOURCES[index] == "Blank"
            output_columns.append(
                {
                    "excel_col": excel_col,
                    "title": STATIC_HEADERS[index],
                    "source": STATIC_SOURCES[index],
                    "rule": STATIC_RULES[index],
                    "result_line": "Blank" if is_blank_source else "Static value in inserted row 2",
                    "section": "Static",
                }
            )

    for projection_row in range(6, 289):
        output_columns.append(
            {
                "excel_col": get_column_letter(len(output_columns) + 1),
                "title": f"Projection Row {projection_row}",
                "source": f"Projections!<qualifying column>{projection_row}",
                "rule": "Copy the cell from the qualifying projection column where row 6 = December and row 8 = Actual.",
                "result_line": f"Dynamic transpose value from Projections row {projection_row}",
                "section": "Dynamic",
            }
        )

    return output_columns


def style_cell(cell, fill_color: str, bold: bool = False, wrap: bool = True) -> None:
    cell.fill = PatternFill("solid", fgColor=fill_color)
    cell.font = Font(bold=bold, color="000000")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
    cell.border = Border(
        left=Side(style="thin", color="666666"),
        right=Side(style="thin", color="666666"),
        top=Side(style="thin", color="666666"),
        bottom=Side(style="thin", color="666666"),
    )


def build_workbook() -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Output Mapping"

    output_columns = build_output_columns()

    ws.freeze_panes = "A6"
    ws.sheet_view.zoomScale = 85

    ws["A1"] = "Valuation Model Output Mapping Template"
    ws["A2"] = "Row 1 shows the output workbook column title. Row 2 shows where the value comes from. Row 3 explains the rule. Row 4 states what the inserted output row should contain."
    ws["A3"] = "Blank title columns are intentional placeholders and should remain blank."

    for ref in ("A1", "A2", "A3"):
        ws[ref].font = Font(bold=True if ref == "A1" else False, size=14 if ref == "A1" else 11)

    row_labels = {
        5: "Output Excel Column",
        6: "Output Title",
        7: "Source",
        8: "Rule",
        9: "Expected Result Line",
        10: "Section",
    }

    for row_idx, label in row_labels.items():
        ws.cell(row=row_idx, column=1).value = label
        style_cell(ws.cell(row=row_idx, column=1), "D9EAF7", bold=True)

    for offset, item in enumerate(output_columns, start=2):
        ws.cell(row=5, column=offset).value = item["excel_col"]
        ws.cell(row=6, column=offset).value = item["title"]
        ws.cell(row=7, column=offset).value = item["source"]
        ws.cell(row=8, column=offset).value = item["rule"]
        ws.cell(row=9, column=offset).value = item["result_line"]
        ws.cell(row=10, column=offset).value = item["section"]

        title_fill = "FCE5CD" if item["section"] == "Static" else "FFF2CC" if item["section"] == "Reserved blank" else "D9EAD3"
        style_cell(ws.cell(row=5, column=offset), "D9EAF7", bold=True)
        style_cell(ws.cell(row=6, column=offset), title_fill, bold=True)
        style_cell(ws.cell(row=7, column=offset), "FFFFFF")
        style_cell(ws.cell(row=8, column=offset), "FFFFFF")
        style_cell(ws.cell(row=9, column=offset), "FFFFFF")
        style_cell(ws.cell(row=10, column=offset), title_fill)

    ws.column_dimensions["A"].width = 24
    for col_idx in range(2, len(output_columns) + 2):
        ws.column_dimensions[get_column_letter(col_idx)].width = 18

    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 42
    ws.row_dimensions[3].height = 24
    ws.row_dimensions[8].height = 44
    ws.row_dimensions[9].height = 32

    legend = wb.create_sheet("Legend")
    legend["A1"] = "Color"
    legend["B1"] = "Meaning"
    style_cell(legend["A1"], "D9EAF7", bold=True)
    style_cell(legend["B1"], "D9EAF7", bold=True)

    legend_rows = [
        ("FCE5CD", "Static columns A:BC populated from Start Page cells, filters, automation values, or intentional blanks."),
        ("D9EAD3", "Dynamic columns BD:LZ populated by transposing Projections rows 6:288."),
    ]
    for row_idx, (color, meaning) in enumerate(legend_rows, start=2):
        legend.cell(row=row_idx, column=1).value = ""
        style_cell(legend.cell(row=row_idx, column=1), color)
        legend.cell(row=row_idx, column=2).value = meaning
        style_cell(legend.cell(row=row_idx, column=2), "FFFFFF")

    legend.column_dimensions["A"].width = 12
    legend.column_dimensions["B"].width = 90

    register = wb.create_sheet("Column Register")
    register_headers = [
        "Output Excel Column",
        "Output Title",
        "Section",
        "Source",
        "Rule",
        "Expected Result Line",
    ]
    for column_idx, title in enumerate(register_headers, start=1):
        register.cell(row=1, column=column_idx).value = title
        style_cell(register.cell(row=1, column=column_idx), "D9EAF7", bold=True)

    for row_idx, item in enumerate(output_columns, start=2):
        values = [
            item["excel_col"],
            item["title"],
            item["section"],
            item["source"],
            item["rule"],
            item["result_line"],
        ]
        fill = "FCE5CD" if item["section"] == "Static" else "FFF2CC" if item["section"] == "Reserved blank" else "D9EAD3"
        for column_idx, value in enumerate(values, start=1):
            register.cell(row=row_idx, column=column_idx).value = value
            style_cell(register.cell(row=row_idx, column=column_idx), fill if column_idx <= 3 else "FFFFFF")

    register.freeze_panes = "A2"
    register.column_dimensions["A"].width = 18
    register.column_dimensions["B"].width = 28
    register.column_dimensions["C"].width = 16
    register.column_dimensions["D"].width = 32
    register.column_dimensions["E"].width = 70
    register.column_dimensions["F"].width = 44

    return wb


def main() -> None:
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    workbook = build_workbook()
    workbook.save(OUTPUT_PATH)
    print(f"Wrote {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
