#!/usr/bin/env python3
"""Generate budget_2026.xlsx — replicates the budget_itemized.xlsx structure.

Each expense category gets its own sheet (Date/Description/Amount, 300 rows).
Summary sheet shows Budget vs Spent per category with live formulas.
VBA Setup sheet includes auto-date macro instructions.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── Categories (matches expense_logger.html) ──
CATEGORIES = [
    ("🍽️", "Eating Out"),
    ("🍻", "Drinking"),
    ("🛍️", "Shopping (want)"),
    ("🛒", "Shopping (need)"),
    ("🥦", "Groceries"),
    ("✈️", "Travel - Flights & Hotel"),
    ("🍜", "Travel - Eat & Drink"),
]

# ── Color palette (dark theme) ──
BG         = PatternFill("solid", fgColor="1E1E2E")
ROW_EVEN   = PatternFill("solid", fgColor="252535")
ROW_ODD    = PatternFill("solid", fgColor="1E1E2C")
HEADER_BG  = PatternFill("solid", fgColor="2A2A3E")

TITLE_FONT   = lambda color: Font(name="Calibri", color=color, size=16, bold=True)
SUBTITLE     = Font(name="Calibri", color="7070A0", size=9)
HDR_FONT     = Font(name="Calibri", color="FFFFFF", size=10, bold=True)
CAT_FONT     = Font(name="Calibri", color="B0B0C8", size=10)
BUDGET_FONT  = Font(name="Calibri", color="4488FF", size=10, bold=True)
SPENT_FONT   = Font(name="Calibri", color="56CFB2", size=10)
REMAIN_FONT  = Font(name="Calibri", color="F7846A", size=10)
TOTAL_FONT   = Font(name="Calibri", color="7C6AF7", size=12, bold=True)
TOTAL_SPENT  = Font(name="Calibri", color="56CFB2", size=12, bold=True)
TOTAL_REMAIN = Font(name="Calibri", color="F7846A", size=12, bold=True)
DATE_FONT    = Font(name="Calibri", color="56CFB2", size=10)
DESC_FONT    = Font(name="Calibri", color="B0B0C8", size=10)
AMT_FONT     = Font(name="Calibri", color="4488FF", size=10)
NOTE_FONT    = Font(name="Calibri", color="56CF72", size=10, bold=True)

MONEY_FMT = '$#,##0.00;($#,##0.00);"-"'
DATE_FMT  = 'MMM DD, YYYY'

DATA_ROWS = 300  # rows 4..303


def fill_bg(ws, max_row, max_col):
    """Fill background for the visible area."""
    for r in range(1, max_row + 2):
        for c in range(1, max_col + 2):
            ws.cell(r, c).fill = BG


def build_category_sheet(wb, emoji, label):
    """Create a per-category expense sheet: Date | Description | Amount."""
    sheet_name = f"{emoji} {label}"
    ws = wb.create_sheet(title=sheet_name)
    fill_bg(ws, DATA_ROWS + 5, 4)

    # Title
    ws["A1"] = sheet_name
    ws["A1"].font = TITLE_FONT("7C6AF7")
    ws["A1"].fill = BG

    # Subtitle
    ws["A2"] = "✏️  Type a description (col B) or amount (col C) — date auto-fills in col A"
    ws["A2"].font = SUBTITLE
    ws["A2"].fill = BG

    # Headers
    for col, header in enumerate(["Date (auto)", "Description", "Amount ($)"], 1):
        cell = ws.cell(row=3, column=col, value=header)
        cell.font = HDR_FONT
        cell.fill = HEADER_BG

    # Data rows with alternating stripes
    for r in range(4, 4 + DATA_ROWS):
        row_fill = ROW_EVEN if (r % 2 == 0) else ROW_ODD
        ws.cell(r, 1).fill = row_fill
        ws.cell(r, 1).font = DATE_FONT
        ws.cell(r, 1).number_format = DATE_FMT

        ws.cell(r, 2).fill = row_fill
        ws.cell(r, 2).font = DESC_FONT

        ws.cell(r, 3).fill = row_fill
        ws.cell(r, 3).font = AMT_FONT
        ws.cell(r, 3).number_format = MONEY_FMT

    # Total row at bottom
    total_row = 4 + DATA_ROWS  # row 304
    ws.cell(total_row, 2).value = "TOTAL"
    ws.cell(total_row, 2).font = TOTAL_FONT
    ws.cell(total_row, 2).fill = BG

    ws.cell(total_row, 3).value = f"=SUM(C4:C{total_row - 1})"
    ws.cell(total_row, 3).font = TOTAL_SPENT
    ws.cell(total_row, 3).fill = BG
    ws.cell(total_row, 3).number_format = MONEY_FMT

    # Column widths
    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 14

    return sheet_name


def build_summary(wb, sheet_names):
    """Build the Summary sheet with Budget / Spent / Remaining per category."""
    ws = wb.create_sheet(title="Summary")
    fill_bg(ws, 20, 5)

    # Title
    ws["A1"] = "💰  Budget Tracker — Summary"
    ws["A1"].font = Font(name="Calibri", color="56CFB2", size=18, bold=True)
    ws["A1"].fill = BG

    # Subtitle
    ws["A2"] = "Type your monthly budget limits in the blue cells below. Spent totals update live."
    ws["A2"].font = SUBTITLE
    ws["A2"].fill = BG

    # Headers
    for col, header in enumerate(["Category", "Monthly Budget", "Spent", "Remaining"], 1):
        cell = ws.cell(row=3, column=col, value=header)
        cell.font = HDR_FONT
        cell.fill = HEADER_BG

    # Category rows
    for i, ((emoji, label), sname) in enumerate(zip(CATEGORIES, sheet_names)):
        row = 4 + i
        row_fill = ROW_EVEN if (row % 2 == 0) else ROW_ODD

        # Category name
        ws.cell(row, 1).value = f"{emoji} {label}"
        ws.cell(row, 1).font = CAT_FONT
        ws.cell(row, 1).fill = row_fill

        # Budget (user editable, blue)
        ws.cell(row, 2).value = 0
        ws.cell(row, 2).font = BUDGET_FONT
        ws.cell(row, 2).fill = row_fill
        ws.cell(row, 2).number_format = MONEY_FMT

        # Spent (pulls SUM from category sheet row 304)
        ws.cell(row, 3).value = f"='{sname}'!C304"
        ws.cell(row, 3).font = SPENT_FONT
        ws.cell(row, 3).fill = row_fill
        ws.cell(row, 3).number_format = MONEY_FMT

        # Remaining
        ws.cell(row, 4).value = f'=IF(B{row}=0,"-",B{row}-C{row})'
        ws.cell(row, 4).font = REMAIN_FONT
        ws.cell(row, 4).fill = row_fill
        ws.cell(row, 4).number_format = MONEY_FMT

    # Total row
    total_row = 4 + len(CATEGORIES)  # row 11
    ws.cell(total_row, 1).value = "TOTAL"
    ws.cell(total_row, 1).font = TOTAL_FONT
    ws.cell(total_row, 1).fill = HEADER_BG

    ws.cell(total_row, 2).value = f"=SUM(B4:B{total_row - 1})"
    ws.cell(total_row, 2).font = TOTAL_FONT
    ws.cell(total_row, 2).fill = HEADER_BG
    ws.cell(total_row, 2).number_format = MONEY_FMT

    ws.cell(total_row, 3).value = f"=SUM(C4:C{total_row - 1})"
    ws.cell(total_row, 3).font = TOTAL_SPENT
    ws.cell(total_row, 3).fill = HEADER_BG
    ws.cell(total_row, 3).number_format = MONEY_FMT

    ws.cell(total_row, 4).value = f'=IF(B{total_row}=0,"-",B{total_row}-C{total_row})'
    ws.cell(total_row, 4).font = TOTAL_REMAIN
    ws.cell(total_row, 4).fill = HEADER_BG
    ws.cell(total_row, 4).number_format = MONEY_FMT

    # Note about VBA
    ws.cell(total_row + 2, 1).value = "⚡  See the 'VBA Setup' tab to enable auto-date (takes 1 minute)"
    ws.cell(total_row + 2, 1).font = NOTE_FONT
    ws.cell(total_row + 2, 1).fill = BG

    # Column widths
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14


def build_vba_setup(wb):
    """Create the VBA Setup instructions sheet."""
    ws = wb.create_sheet(title="VBA Setup")
    fill_bg(ws, 50, 3)

    lines = [
        (1, "⚡  Enable Auto-Date (one-time setup)", Font(name="Calibri", color="56CFB2", size=16, bold=True)),
        (3, "HOW TO ENABLE AUTO-DATE IN EXCEL (Windows)", Font(name="Calibri", color="7C6AF7", size=12, bold=True)),
        (5, "Step 1 → Press  Alt + F11  to open the Visual Basic Editor", Font(name="Calibri", color="B0B0C8", size=10)),
        (6, "Step 2 → In the left panel, double-click 'ThisWorkbook'", Font(name="Calibri", color="B0B0C8", size=10)),
        (7, "Step 3 → Copy ALL the code below and paste it into the editor", Font(name="Calibri", color="B0B0C8", size=10)),
        (8, "Step 4 → Press  Ctrl + S  (save as .xlsm if prompted)", Font(name="Calibri", color="B0B0C8", size=10)),
        (9, "Step 5 → Close the VBA editor. Done! ✅", Font(name="Calibri", color="B0B0C8", size=10)),
        (11, "HOW TO ENABLE AUTO-DATE IN EXCEL (Mac)", Font(name="Calibri", color="7C6AF7", size=12, bold=True)),
        (13, "Step 1 → Go to Tools > Macros > Edit Macros", Font(name="Calibri", color="B0B0C8", size=10)),
        (14, "Step 2 → Double-click 'ThisWorkbook' in the left panel", Font(name="Calibri", color="B0B0C8", size=10)),
        (15, "Step 3 → Copy ALL the code below and paste it", Font(name="Calibri", color="B0B0C8", size=10)),
        (16, "Step 4 → Save. Done! ✅", Font(name="Calibri", color="B0B0C8", size=10)),
        (18, "NOTE: When opening the file, click 'Enable Macros' if prompted.", Font(name="Calibri", color="F7846A", size=10, bold=True)),
        (20, "━━━  COPY THE CODE BELOW  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", Font(name="Calibri", color="56CFB2", size=10)),
    ]

    for row, text, font in lines:
        ws.cell(row, 1).value = text
        ws.cell(row, 1).font = font
        ws.cell(row, 1).fill = BG

    # VBA code
    vba_code = [
        "' ============================================================",
        "' PASTE THIS INTO: Tools > Macro > Basic IDE > ThisWorkbook",
        "' ============================================================",
        "",
        "Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)",
        "    Dim cell As Range",
        '    If Sh.Name = "Summary" Or Sh.Name = "VBA Setup" Then Exit Sub',
        "    If Target.Column < 2 Or Target.Column > 3 Then Exit Sub",
        "    Application.EnableEvents = False",
        "    On Error GoTo Cleanup",
        "    For Each cell In Target",
        "        If cell.Column >= 2 And cell.Column <= 3 Then",
        "            Dim dateCell As Range",
        "            Set dateCell = Sh.Cells(cell.Row, 1)",
        '            If cell.Value <> "" And dateCell.Value = "" Then',
        "                dateCell.Value = Date",
        '                dateCell.NumberFormat = "MMM DD, YYYY"',
        "            End If",
        '            If Sh.Cells(cell.Row, 2).Value = "" And _',
        '               Sh.Cells(cell.Row, 3).Value = "" Then',
        "                dateCell.ClearContents",
        "            End If",
        "        End If",
        "    Next cell",
        "Cleanup:",
        "    Application.EnableEvents = True",
        "End Sub",
    ]

    code_font = Font(name="Consolas", color="56CF72", size=10)
    for i, line in enumerate(vba_code):
        cell = ws.cell(row=21 + i, column=1, value=line)
        cell.font = code_font
        cell.fill = BG

    ws.column_dimensions["A"].width = 80


def main():
    wb = Workbook()
    wb.remove(wb.active)

    # Build category sheets
    sheet_names = []
    for emoji, label in CATEGORIES:
        sname = build_category_sheet(wb, emoji, label)
        sheet_names.append(sname)

    # Build summary (references category sheets)
    build_summary(wb, sheet_names)

    # Build VBA setup
    build_vba_setup(wb)

    # Reorder: Summary first, VBA Setup last
    summary_idx = wb.sheetnames.index("Summary")
    wb.move_sheet("Summary", offset=-summary_idx)

    out = "budget_2026.xlsx"
    wb.save(out)
    print(f"✓ Created {out}")
    print(f"  Sheets: {', '.join(wb.sheetnames)}")
    print(f"  Each category has {DATA_ROWS} entry rows with auto-sum")
    print(f"  Summary links to all category totals with Budget/Spent/Remaining")


if __name__ == "__main__":
    main()
