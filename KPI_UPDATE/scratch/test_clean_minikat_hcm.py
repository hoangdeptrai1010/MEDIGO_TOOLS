# -*- coding: utf-8 -*-
import openpyxl, os, sys
sys.stdout.reconfigure(encoding='utf-8')

# HCM 9 Branches and their official staff (matching BẢNG LƯƠNG exactly)
hcm_branches = {
    'Trường Sa': [
        'Nguyễn Trần Ngọc Phương',
        'Nguyễn Ngọc Anh Thư',
        'Ngô Trần Tuyết Vy',
        'Phạm Nguyễn Ngọc Quý',
        'Trần Thị Kim Khánh'
    ],
    'Đỗ Quang Đẩu': [
        'Ngô Thị Thanh Thắm',
        'Phạm Thị Nghĩa Hương',
        'Trần Thiên Phát',
        'Hoàng Thanh Thủy',
        'Lê Thị Huyền Trân'
    ],
    'Minh Châu': [
        'Thái Thùy Linh',
        'Trịnh Thị Phượng',
        'Phan Công Vũ Tài',
        'Đoàn Quốc Tâm',
        'Nguyễn Ngọc Thùy'
    ],
    'Nam Hòa': [
        'Dương Thị Huỳnh Như',
        'Trần Thị Ánh Nguyệt',
        'Đỗ Thị Kim Tiến',
        'Bùi Thị Thanh Thủy',
        'Nguyễn Thị Huyền Trang',
        'Trần Linh Khải'
    ],
    'Nguyễn Chí Thanh': [
        'Lê Thị Quỳnh Trâm',
        'Nguyễn Thị Hồng Hạnh',
        'Trần Hoàng Khánh',
        'Lý Thục Mi',
        'Đỗ Thị Phương Thảo'
    ],
    'Nguyễn Thị Thập': [
        'Triệu Thị Ngọc Lý',
        'Cao Trọng Nhân',
        'Trần Thành Đạt',
        'Nguyễn Trí Nghĩa',
        'Bích Trâm (Đã nghỉ)',
        'Bùi Thị Bích Trâm'
    ],
    'Nguyễn Văn Quá': [
        'Võ Ngọc Giàu Sang',
        'Hồ Ngọc Lý',
        'Nguyễn Thị Thu Huyền',
        'Nguyễn Thị Hương Giang'
    ],
    'Rạch Bùng Binh': [
        'Cù Thị Tường Vy',
        'Huỳnh Thị Huệ Hiền',
        'Nguyễn Thị Phúc Lộc',
        'Ngô Thị Ngọc Thủy',
        'Phạm Nguyễn Ngọc Quyền Trân'
    ],
    'Lê Bình': [
        'Hồ Thị Minh Hòa',
        'Lê Thị Xuyến',
        'Đinh Hoài Bảo',
        'Hoàng Lâm Gia Bảo',
        'Trần Thị Phương Thùy'
    ]
}

print(f"Total HCM Branches: {len(hcm_branches)}")
tot_staff = sum(len(staffs) for staffs in hcm_branches.values())
print(f"Total HCM Staff: {tot_staff}")
