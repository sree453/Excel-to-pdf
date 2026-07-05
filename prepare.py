"""
prepare.py — normalizes a spreadsheet so it prints as a clean table.

Usage:
    python3 prepare.py csv   <input.csv>   <output.xlsx>
    python3 prepare.py style <input.xlsx>  <output.xlsx>

"csv" mode reads a CSV and writes a styled workbook.
"style" mode opens an existing workbook and applies the same table
styling + print setup, without touching the cell values.
"""

import sys
import os
import csv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


def style_workbook(wb):
    thin = Side(style='thin', color='999999')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill(start_color='E8EFEA', end_color='E8EFEA', fill_type='solid')
    header_font = Font(bold=True)

    for ws in wb.worksheets:
        max_row = ws.max_row
        max_col = ws.max_column
        if max_row == 0 or max_col == 0:
            continue

        # Real borders on every used cell, so the PDF always shows gridlines
        # (relying on the "print gridlines" checkbox is inconsistent across apps).
        for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
            for cell in row:
                cell.border = border

        # Bold, shaded header row.
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font

        # Rough auto-width so columns aren't cramped or absurdly wide.
        for col_idx in range(1, max_col + 1):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            scan_rows = min(max_row, 300)
            for row_idx in range(1, scan_rows + 1):
                val = ws.cell(row=row_idx, column=col_idx).value
                if val is not None:
                    max_len = max(max_len, len(str(val)))
            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 8), 40)

        # Page setup: force every column onto one page width, let rows
        # flow across as many pages tall as needed, and repeat the header.
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_title_rows = '1:1'

        last_col_letter = get_column_letter(max_col)
        ws.print_area = f"A1:{last_col_letter}{max_row}"


def convert_csv_to_xlsx(csv_path, xlsx_path):
    wb = Workbook()
    ws = wb.active
    with open(csv_path, newline='', encoding='utf-8-sig', errors='replace') as f:
        reader = csv.reader(f)
        for row in reader:
            ws.append(row)
    style_workbook(wb)
    wb.save(xlsx_path)


def restyle_xlsx(input_path, output_path):
    wb = load_workbook(input_path)
    style_workbook(wb)
    wb.save(output_path)


def main():
    if len(sys.argv) != 4:
        print('Usage: prepare.py <csv|style> <input> <output>', file=sys.stderr)
        sys.exit(1)

    mode, input_path, output_path = sys.argv[1], sys.argv[2], sys.argv[3]

    if mode == 'csv':
        convert_csv_to_xlsx(input_path, output_path)
    elif mode == 'style':
        restyle_xlsx(input_path, output_path)
    else:
        print(f'Unknown mode: {mode}', file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
