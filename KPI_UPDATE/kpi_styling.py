import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import unicodedata
import re

# Fonts
font_header_white = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
font_header_dark = Font(name='Calibri', size=11, bold=True, color='1E293B')
font_header_green = Font(name='Calibri', size=11, bold=True, color='166534')
font_bold = Font(name='Calibri', size=11, bold=True)
font_regular = Font(name='Calibri', size=11)
font_bonus = Font(name='Calibri', size=11, bold=True, color='15803D')

# Header Fills
fill_navy = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')      # Primary Navy
fill_soft_blue = PatternFill(start_color='DDEBF7', end_color='DDEBF7', fill_type='solid') # Soft Blue
fill_light_blue = PatternFill(start_color='BDD7EE', end_color='BDD7EE', fill_type='solid')# Light Blue
fill_yellow = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')    # KPI / Highlight Yellow
fill_green = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')     # Bonus / Success Green
fill_emerald = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')   # Strong Green
fill_orange = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')    # Target / Hàng điểm Orange
fill_purple = PatternFill(start_color='E8D8F8', end_color='E8D8F8', fill_type='solid')    # Project Purple
fill_gray = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')      # Subtle Gray
fill_dark_gray = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid') # Dark Gray

# Cell highlight fills
fill_highlight_yellow = PatternFill(start_color='FFFBEA', end_color='FFFBEA', fill_type='solid')
fill_highlight_mint = PatternFill(start_color='EBF9F1', end_color='EBF9F1', fill_type='solid')
fill_row_zebra = PatternFill(start_color='FAFAFA', end_color='FAFAFA', fill_type='solid')

# Borders
border_thin = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)
border_header = Border(
    left=Side(style='thin', color='B0C4DE'),
    right=Side(style='thin', color='B0C4DE'),
    top=Side(style='medium', color='1F4E79'),
    bottom=Side(style='medium', color='1F4E79')
)

def strip_accents_lower(s):
    if not s:
        return ""
    nfkd = unicodedata.normalize('NFKD', str(s))
    res = "".join([c for c in nfkd if not unicodedata.combining(c)])
    return res.lower().replace('đ', 'd').replace('Đ', 'd').strip()

def apply_full_kpi_styling(wb):
    """
    Applies unified, beautiful pharmacy color styling across all KPI workbook sheets:
    - 'data'
    - 'kpi dược sĩ'
    - 'kpi nhà thuốc'
    - 'Dự án T...'
    - 'Hot Bill HN'
    """
    for sname in wb.sheetnames:
        ws = wb[sname]
        try:
            ws.views.sheetView[0].showGridLines = True
        except Exception:
            pass

        norm_name = strip_accents_lower(sname)

        # -------------------------------------------------------------
        # 1. SHEET: 'data'
        # -------------------------------------------------------------
        if norm_name == 'data':
            for c_i in range(1, min(ws.max_column + 1, 12)):
                cell = ws.cell(1, c_i)
                if c_i in (1, 2):
                    cell.font = font_header_white
                    cell.fill = fill_navy
                elif c_i in (3, 4):
                    cell.font = font_header_dark
                    cell.fill = fill_soft_blue
                elif c_i in (5, 6):
                    cell.font = font_header_dark
                    cell.fill = fill_orange
                elif c_i in (7, 8):
                    cell.font = font_header_green
                    cell.fill = fill_green
                else:
                    cell.font = font_header_dark
                    cell.fill = fill_gray
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border_header

            for r in range(2, ws.max_row + 1):
                if not ws.cell(r, 1).value and not ws.cell(r, 2).value:
                    continue
                is_even = (r % 2 == 0)
                row_bg = fill_row_zebra if is_even else None
                for c_i in range(1, min(ws.max_column + 1, 12)):
                    cell = ws.cell(r, c_i)
                    cell.font = font_regular
                    cell.border = border_thin
                    if row_bg:
                        cell.fill = row_bg
                    if c_i in (4, 6, 8):
                        cell.number_format = '#,##0'
                    if c_i in (1, 2):
                        cell.alignment = Alignment(horizontal='left')
                    elif c_i in (3, 5, 7):
                        cell.alignment = Alignment(horizontal='center')
                    else:
                        cell.alignment = Alignment(horizontal='right')

        # -------------------------------------------------------------
        # 2. SHEET: 'kpi dược sĩ' / 'kpi duoc si'
        # -------------------------------------------------------------
        elif 'duoc si' in norm_name or 'ds' in norm_name:
            # Row 1
            ws.cell(1, 1).font = font_bold
            ws.cell(1, 3).font = font_bold
            ws.cell(1, 3).number_format = 'dd/mm/yyyy'

            # Row 2 (Headers)
            for c_i in range(1, min(ws.max_column + 1, 35)):
                cell = ws.cell(2, c_i)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.border = border_header
                if c_i in (1, 2, 3):
                    cell.font = font_header_white
                    cell.fill = fill_navy
                elif c_i in (4, 5):
                    cell.font = font_header_dark
                    cell.fill = fill_soft_blue
                elif c_i in (6, 7):
                    cell.font = font_header_dark
                    cell.fill = fill_purple
                elif c_i in (8, 9, 10):
                    cell.font = font_header_dark
                    cell.fill = fill_yellow
                elif c_i in (11, 12, 13, 14):
                    cell.font = font_header_green
                    cell.fill = fill_emerald
                elif c_i in range(15, 21):
                    cell.font = font_header_dark
                    cell.fill = fill_dark_gray
                    cell.number_format = '0%'
                elif c_i in (21, 22):
                    cell.font = font_header_dark
                    cell.fill = fill_light_blue
                else:
                    cell.font = font_header_dark
                    cell.fill = fill_gray

            # Data rows
            for r in range(3, ws.max_row + 1):
                if not ws.cell(r, 2).value:
                    continue
                is_even = (r % 2 == 0)
                row_tint = fill_row_zebra if is_even else None
                for c_i in range(1, min(ws.max_column + 1, 35)):
                    cell = ws.cell(r, c_i)
                    cell.font = font_regular
                    cell.border = border_thin
                    if row_tint:
                        cell.fill = row_tint

                    if c_i == 10: # % KPI
                        cell.fill = fill_highlight_yellow
                        cell.font = font_bold
                        cell.number_format = '0.0%'
                    elif c_i in (11, 13, 14): # Tiền thưởng
                        if c_i == 14:
                            cell.fill = fill_highlight_mint
                            cell.font = font_bonus
                        cell.number_format = '#,##0'
                    elif c_i in (4, 5, 6, 7, 8, 9, 15, 16, 17, 18, 19, 20, 21, 22, 24, 25, 27, 28, 30, 31):
                        cell.number_format = '#,##0'

                    if c_i in (1, 2, 3):
                        cell.alignment = Alignment(horizontal='left')
                    elif c_i in (10, 12, 15, 16, 17, 18, 19, 20, 23, 26, 29):
                        cell.alignment = Alignment(horizontal='center')
                    else:
                        cell.alignment = Alignment(horizontal='right')

        # -------------------------------------------------------------
        # 3. SHEET: 'kpi nhà thuốc' / 'kpi nha thuoc'
        # -------------------------------------------------------------
        elif 'nha thuoc' in norm_name or 'nt' in norm_name:
            # Row 1
            ws.cell(1, 4).font = font_bold
            ws.cell(1, 6).number_format = 'dd/mm/yyyy'
            ws.cell(1, 20).number_format = 'dd/mm/yyyy'

            # Row 2 (Headers)
            for c_i in range(1, min(ws.max_column + 1, 35)):
                cell = ws.cell(2, c_i)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.border = border_header
                if c_i in (1, 2):
                    cell.font = font_header_white
                    cell.fill = fill_navy
                elif c_i in (3, 4, 5, 6):
                    cell.font = font_header_dark
                    cell.fill = fill_soft_blue
                elif c_i == 7:
                    cell.font = font_header_green
                    cell.fill = fill_emerald
                elif c_i in range(8, 21):
                    cell.font = font_header_dark
                    cell.fill = fill_yellow if c_i in (8, 13, 17) else fill_gray
                elif c_i in (21, 22):
                    cell.font = font_header_dark
                    cell.fill = fill_orange
                elif c_i == 23:
                    cell.font = font_header_green
                    cell.fill = fill_emerald
                else:
                    cell.font = font_header_dark
                    cell.fill = fill_gray

            # Data rows
            for r in range(3, ws.max_row + 1):
                if not ws.cell(r, 2).value:
                    continue
                is_even = (r % 2 == 0)
                row_tint = fill_row_zebra if is_even else None
                for c_i in range(1, min(ws.max_column + 1, 35)):
                    cell = ws.cell(r, c_i)
                    cell.font = font_regular
                    cell.border = border_thin
                    if row_tint:
                        cell.fill = row_tint

                    if c_i == 7: # Thưởng CHT
                        cell.fill = fill_highlight_mint
                        cell.font = font_bonus
                        cell.number_format = '#,##0'
                    elif c_i == 23: # Xét Đạt
                        cell.font = font_bold
                        cell.alignment = Alignment(horizontal='center')
                    elif c_i in (3, 4, 5, 9, 10, 11, 12, 14, 15, 16, 18, 19, 20, 21, 25, 26, 28, 29, 31, 32):
                        cell.number_format = '#,##0'
                    elif c_i in (6, 8, 13, 17, 22):
                        cell.number_format = '0.0%'

                    if c_i in (1, 2):
                        cell.alignment = Alignment(horizontal='left')
                    elif c_i in (6, 8, 13, 17, 22, 23, 24, 27, 30):
                        cell.alignment = Alignment(horizontal='center')
                    else:
                        cell.alignment = Alignment(horizontal='right')

        # -------------------------------------------------------------
        # 4. SHEET: 'Dự án...' / 'du an'
        # -------------------------------------------------------------
        elif 'du an' in norm_name or 'project' in norm_name:
            ws.cell(1, 1).font = font_bold
            ws.cell(1, 1).number_format = 'dd/mm/yyyy'

            # Row 2 (Headers)
            for c_i in range(1, min(ws.max_column + 1, 15)):
                cell = ws.cell(2, c_i)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border_header
                if c_i in (1, 2, 3):
                    cell.font = font_header_white
                    cell.fill = fill_navy
                elif c_i in (4, 5, 6, 7, 8, 9):
                    cell.font = font_header_dark
                    cell.fill = fill_purple
                elif c_i in (10, 11, 12, 13):
                    cell.font = font_header_green
                    cell.fill = fill_emerald
                else:
                    cell.font = font_header_dark
                    cell.fill = fill_gray

            # Data rows
            for r in range(3, ws.max_row + 1):
                if not ws.cell(r, 2).value:
                    continue
                is_even = (r % 2 == 0)
                row_tint = fill_row_zebra if is_even else None
                for c_i in range(1, min(ws.max_column + 1, 15)):
                    cell = ws.cell(r, c_i)
                    cell.font = font_regular
                    cell.border = border_thin
                    if row_tint:
                        cell.fill = row_tint

                    if c_i == 13: # Total dự án
                        cell.fill = fill_highlight_mint
                        cell.font = font_bonus
                        cell.number_format = '#,##0'
                    elif c_i in (4, 5, 6, 7, 8, 9, 10, 11, 12):
                        cell.number_format = '#,##0'

                    if c_i in (1, 2, 3):
                        cell.alignment = Alignment(horizontal='left')
                    else:
                        cell.alignment = Alignment(horizontal='right')

        # -------------------------------------------------------------
        # 5. SHEET: 'Hot Bill HN' / 'hot bill'
        # -------------------------------------------------------------
        elif 'hot bill' in norm_name:
            for c_i in range(1, min(ws.max_column + 1, 10)):
                cell = ws.cell(1, c_i)
                if c_i in (1, 2, 3, 4):
                    cell.font = font_header_white
                    cell.fill = fill_navy
                elif c_i in (5, 6):
                    cell.font = font_header_dark
                    cell.fill = fill_soft_blue
                elif c_i == 7:
                    cell.font = font_header_green
                    cell.fill = fill_emerald
                else:
                    cell.font = font_header_dark
                    cell.fill = fill_gray
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border_header

            for r in range(2, ws.max_row + 1):
                if not ws.cell(r, 1).value and not ws.cell(r, 4).value:
                    continue
                is_even = (r % 2 == 0)
                row_tint = fill_row_zebra if is_even else None
                for c_i in range(1, min(ws.max_column + 1, 10)):
                    cell = ws.cell(r, c_i)
                    cell.font = font_regular
                    cell.border = border_thin
                    if row_tint:
                        cell.fill = row_tint
                    if c_i in (5, 6, 7):
                        cell.number_format = '#,##0'
                    if c_i == 7:
                        cell.font = font_bonus

        # Auto-fit Column Widths with comfortable padding
        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = 0
            for cell in col[:20]:
                val = str(cell.value or '')
                if len(val) > max_len:
                    max_len = len(val)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)
