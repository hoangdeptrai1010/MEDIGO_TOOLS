# -*- coding: utf-8 -*-
"""
payroll_minikat_writer.py
==========================
Ghi sheet "MiniProjects" vào workbook bảng lương — SỐ CỘT SINH RA TỰ ĐỘNG
theo số dự án mini trong config (3 hay 7 dự án đều ra đúng, không cần sửa code).

Khác với sheet "MiniKat-HN"/"MiniKat-HCM" cũ (cột cố định, chỉ chứa đúng 4
dự án đã biết trước tên), sheet này lặp qua danh sách dự án trong config và
tự tính vị trí cột — thêm dự án thứ 5, 6, 7 vào config là sheet tự giãn ra
thêm cột, không đụng vào file này.

Mỗi dự án chiếm 3 cột: [Số lượng/Doanh thu đạt] [Thưởng cá nhân] [Đủ điều kiện thưởng NT?]
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from mini_projects_core import MiniProjectsEngine

HEADER_FONT = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
HEADER_FILL = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
PROJECT_FILL_ALT = PatternFill(start_color='EDF2F8', end_color='EDF2F8', fill_type='solid')
THIN = Side(style='thin', color='D9D9D9')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def write_minikat_sheet(wb_out, engine: MiniProjectsEngine, sellers_with_branch, sheet_name="MiniProjects"):
    """
    sellers_with_branch: list các tuple (seller_name, branch) cần liệt kê ra sheet
                          (thường lấy từ toàn bộ nhân sự trong bảng lương, không chỉ
                          người có phát sinh dự án mini, để không "mất" ai khỏi sheet).
    """
    if sheet_name in wb_out.sheetnames:
        del wb_out[sheet_name]
    ws = wb_out.create_sheet(sheet_name)

    # ---- Dòng tiêu đề: cột A,B cố định (Chi nhánh, Tên NV), sau đó lặp theo dự án ----
    ws.cell(1, 1, "Chi nhánh").font = HEADER_FONT
    ws.cell(1, 2, "Tên nhân viên").font = HEADER_FONT
    ws.cell(1, 1).fill = HEADER_FILL
    ws.cell(1, 2).fill = HEADER_FILL

    col = 3
    project_col_start = {}
    for proj in engine.projects:
        project_col_start[proj.id] = col
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 1)
        c = ws.cell(1, col, proj.name)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal='center')
        ws.cell(2, col, "SL / Doanh thu").font = Font(bold=True, size=9)
        ws.cell(2, col + 1, "Thưởng cá nhân").font = Font(bold=True, size=9)
        col += 2

    total_col = col
    ws.cell(1, total_col, "Tổng thưởng dự án mini").font = HEADER_FONT
    ws.cell(1, total_col).fill = HEADER_FILL
    ws.merge_cells(start_row=1, start_column=total_col, end_row=2, end_column=total_col)

    # ---- Dữ liệu từng người ----
    r = 3
    for seller, branch in sellers_with_branch:
        ws.cell(r, 1, branch)
        ws.cell(r, 2, seller)
        row_total = 0.0
        for proj in engine.projects:
            c0 = project_col_start[proj.id]
            qty, rev, _distinct = engine.get_seller_value(seller, proj.id)
            region = proj.regions.get(engine.region_of(branch))
            display_val = qty if (region and region.individual_tiers and region.individual_tiers[0].gate_on == "qty") else rev
            bonus = engine.calc_individual_bonus(proj.id, seller, branch) if region else 0.0
            ws.cell(r, c0, round(display_val, 0)).number_format = '#,##0'
            ws.cell(r, c0 + 1, round(bonus, 0)).number_format = '#,##0'
            row_total += bonus
        ws.cell(r, total_col, round(row_total, 0)).number_format = '#,##0'
        for cc in range(1, total_col + 1):
            ws.cell(r, cc).border = BORDER
        r += 1

    # ---- Khối thưởng cửa hàng (store bonus) cho từng dự án — nằm dưới bảng cá nhân ----
    r += 2
    ws.cell(r, 1, "THƯỞNG CỬA HÀNG THEO DỰ ÁN MINI").font = Font(bold=True, size=12)
    r += 1
    ws.cell(r, 1, "Chi nhánh").font = HEADER_FONT
    ws.cell(r, 1).fill = HEADER_FILL
    hdr_row = r
    col = 2
    store_col_start = {}
    for proj in engine.projects:
        store_col_start[proj.id] = col
        c = ws.cell(hdr_row, col, proj.name)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        col += 1
    r += 1

    all_branches = sorted({b for _s, b in sellers_with_branch})
    for branch in all_branches:
        ws.cell(r, 1, branch)
        for proj in engine.projects:
            c0 = store_col_start[proj.id]
            region = proj.regions.get(engine.region_of(branch))
            store_bonus = engine.calc_store_bonus(proj.id, branch) if region else 0.0
            ws.cell(r, c0, round(store_bonus, 0)).number_format = '#,##0'
        r += 1

    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 24
    return ws
