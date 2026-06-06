import csv
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter

# --- Load CSV ---
rows = []
with open("corporate_characters_100.csv", encoding="utf-8") as f:
    reader = csv.reader(f)
    for row in reader:
        rows.append(row)

headers = rows[0]
data = rows[1:]

wb = Workbook()
ws = wb.active
ws.title = "企業キャラクター100事例"

# --- Styles ---
HEADER_FILL   = PatternFill("solid", fgColor="1F3864")   # dark navy
HEADER_FONT   = Font(name="Meiryo UI", bold=True, color="FFFFFF", size=10)
ROW_FILL_ODD  = PatternFill("solid", fgColor="EBF3FB")   # light blue
ROW_FILL_EVEN = PatternFill("solid", fgColor="FFFFFF")
BODY_FONT     = Font(name="Meiryo UI", size=9)
WRAP_ALIGN    = Alignment(wrap_text=True, vertical="top")
CENTER_ALIGN  = Alignment(wrap_text=False, vertical="center", horizontal="center")

thin = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

# --- Column widths (chars) ---
COL_WIDTHS = {
    1:  5,   # No.
    2:  20,  # キャラクター名
    3:  22,  # 企業/ブランド名
    4:  8,   # 国/地域
    5:  16,  # 業種/業界
    6:  8,   # 誕生年
    7:  18,  # キャラクタータイプ
    8:  36,  # 設定・ビジュアル特徴
    9:  28,  # 役割・用途
    10: 18,  # ターゲット層
    11: 28,  # IP展開・グッズ化
    12: 44,  # 成功要因・特記事項
    13: 12,  # 現在の状況
}

# --- Write header row ---
ws.row_dimensions[1].height = 28
for col_idx, header in enumerate(headers, start=1):
    cell = ws.cell(row=1, column=col_idx, value=header)
    cell.fill   = HEADER_FILL
    cell.font   = HEADER_FONT
    cell.border = BORDER
    cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.column_dimensions[get_column_letter(col_idx)].width = COL_WIDTHS.get(col_idx, 15)

# --- Write data rows ---
for row_idx, row in enumerate(data, start=2):
    fill = ROW_FILL_ODD if row_idx % 2 == 0 else ROW_FILL_EVEN
    ws.row_dimensions[row_idx].height = 54

    for col_idx, value in enumerate(row, start=1):
        cell = ws.cell(row=row_idx, column=col_idx, value=value)
        cell.fill   = fill
        cell.font   = BODY_FONT
        cell.border = BORDER

        # No. column → center
        if col_idx == 1:
            cell.alignment = CENTER_ALIGN
        # 国/地域・誕生年・状況 → center
        elif col_idx in (4, 6, 13):
            cell.alignment = Alignment(wrap_text=False, vertical="top", horizontal="center")
        else:
            cell.alignment = WRAP_ALIGN

# --- Freeze header row ---
ws.freeze_panes = "A2"

# --- Auto-filter ---
ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

# --- Sheet tab color ---
ws.sheet_properties.tabColor = "1F3864"

wb.save("corporate_characters_100.xlsx")
print("Saved: corporate_characters_100.xlsx")
