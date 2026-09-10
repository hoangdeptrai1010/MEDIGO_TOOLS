"""
TOOL V5: CÔNG CỤ THEO DÕI, ĐỊNH GIÁ & XUẤT BÁO CÁO GIÁ TUẦN MEDIGO
===================================================================
Quy tắc V5:
  1. Giá vốn gần nhất: Lấy ĐƠN GIÁ SAU VAT CHUẨN = Giá nhập + (VAT nhập hàng / Số lượng), ưu tiên theo thời gian mới nhất, ưu tiên 'Kho Tổng CNT' (giá nhập min), sau đó đến các chi nhánh khác.
  2. Giá vốn cũ: Lấy từ kỳ gần nhất trước đó trong lịch sử 22 kỳ (Sheet 'update giá nhập' của master data) hoặc state tuần trước.
  3. Ma trận định giá 2 tầng:
     - Tầng 1 (Có giá đối thủ):
       + Nếu Margin tại giá đối thủ >= Target Margin ngành (Thuốc 14%, TPCN 15%, TBYT 20%) -> Bán đúng bằng giá đối thủ.
       + Nếu Margin < Target Margin (hoặc Giá vốn >= Đối thủ) -> Cộng thêm % Range đối thủ (+10%, +8%, +6%, +4%, +2.5%) và làm tròn lên hàng nghìn.
     - Tầng 2 (Không có giá đối thủ / Offline Low / Offline High): Tính theo Bảng Target Margin C-F theo từng ngành hàng.
  4. Tự động gom riêng các mã 'Giá vốn Medigo > Giá bán Đối thủ' vào Sheet 'List sp gia nhap cao'.
  5. Tự động loại bỏ các sản phẩm ngừng kinh doanh: (Tồn kho <= 0 & Không có PO & Có giá bán) sang Sheet 'DS bi loai'.
  6. Xuất đồng thời 2 file báo cáo Excel (Bản Online-Only để trống giá Offline cho Sếp duyệt & Bản Full đầy đủ giá gợi ý).
"""

import os
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


# ======================================================================
# 1. CẤU HÌNH ĐƯỜNG DẪN ĐỘNG (CONFIG)
# ======================================================================
SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent

DATA_DIR = PROJECT_DIR / "DATA"
OUTPUT_DIR = PROJECT_DIR / "OUTPUT"
HISTORY_DIR = SCRIPT_DIR / "history"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
HISTORY_DIR.mkdir(parents=True, exist_ok=True)

STATE_FILE = SCRIPT_DIR / "state_gia_tuan_truoc_v5.csv"


def find_latest_file(directory: Path, patterns: list[str], exclude_keywords: list[str] = None, fallback_name: Optional[str] = None) -> Path:
    candidates = []
    for pat in patterns:
        for f in directory.glob(pat):
            if f.name.startswith("~$"):
                continue
            if exclude_keywords and any(k.lower() in f.name.lower() for k in exclude_keywords):
                continue
            candidates.append(f)
    if candidates:
        candidates.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        return candidates[0]
    if fallback_name:
        fb = directory / fallback_name
        if fb.exists() and not fb.name.startswith("~$"):
            return fb
    raise FileNotFoundError(f"Không tìm thấy file nào khớp với patterns {patterns} trong {directory}")


def read_excel_fast(path: Path, sheet_name: any = 0) -> pd.DataFrame:
    try:
        return pd.read_excel(path, sheet_name=sheet_name, engine="calamine")
    except Exception:
        return pd.read_excel(path, sheet_name=sheet_name)


# ======================================================================
# 2. CHUẨN HÓA ĐƠN VỊ TÍNH (DVT NORMALIZER)
# ======================================================================
DVT_MAP = {
    "hop": "hộp", "h": "hộp", "box": "hộp", "hộp": "hộp", "hop ": "hộp",
    "vien": "viên", "v": "viên", "pill": "viên", "tab": "viên", "tablet": "viên", "viên": "viên",
    "vi": "vỉ", "blister": "vỉ", "vỉ": "vỉ",
    "tuyp": "tuýp", "t": "tuýp", "tube": "tuýp", "tuýp": "tuýp",
    "chai": "chai", "c": "chai", "bottle": "chai", "lo": "chai", "lọ": "chai", "hũ": "chai", "hu": "chai",
    "goi": "gói", "g": "gói", "sachet": "gói", "pack": "gói", "tui": "gói", "túi": "gói", "gói": "gói",
    "ong": "ống", "amp": "ống", "ampoule": "ống", "ống": "ống",
    "mieng": "miếng", "patch": "miếng", "miếng": "miếng",
    "binh": "bình", "bình": "bình",
    "cap": "cặp", "bo": "cặp", "cặp": "cặp", "bộ": "cặp",
    "cai": "cái", "cây": "cái", "chiec": "cái", "chiếc": "cái", "cái": "cái",
    "loc": "lốc", "lốc": "lốc", "thung": "thùng", "thùng": "thùng",
}


def norm_str(s: any) -> str:
    if pd.isna(s):
        return ""
    s = str(s).strip().lower()
    return re.sub(r"\s+", " ", s)


def norm_dvt(dvt: any) -> str:
    cleaned = norm_str(dvt)
    plain = (
        cleaned.replace("ộ", "o").replace("ố", "o").replace("ồ", "o").replace("ọ", "o")
        .replace("ê", "e").replace("ế", "e").replace("ề", "e").replace("ệ", "e")
        .replace("í", "i").replace("ỉ", "i").replace("ị", "i")
        .replace("ý", "y").replace("ỳ", "y").replace("ỵ", "y")
        .replace("ú", "u").replace("ù", "u").replace("ụ", "u")
        .replace("ả", "a").replace("ã", "a").replace("á", "a").replace("à", "a").replace("ạ", "a")
        .replace("đ", "d")
    )
    return DVT_MAP.get(cleaned, DVT_MAP.get(plain, cleaned))


def norm_key(*parts) -> str:
    s = "".join(str(p) for p in parts if pd.notna(p))
    return re.sub(r"\s+", "", s.strip().lower())


# ======================================================================
# 3. TẢI DỮ LIỆU & LỊCH SỬ GIÁ VỐN SAU VAT (COGS WATERFALL - ƯU TIÊN KHO TỔNG)
# ======================================================================
def load_lc_mapping(path: Path, app_file_path: Optional[Path] = None) -> pd.DataFrame:
    df = read_excel_fast(path, sheet_name="data_LC")
    keep = {
        "Mã Hàng": "Ma_Hang",
        "Mã sp": "Ma_sp",
        "Tên Hàng": "Ten_Hang",
        "ĐVT": "DVT",
        "Link LC": "Link_LC",
        "BG LC": "BG_LC_ghi_nhan_truoc",
        "Quy Đổi": "Quy_Doi",
    }
    available_keep = {k: v for k, v in keep.items() if k in df.columns}
    df = df[list(available_keep.keys())].rename(columns=available_keep)
    df = df.dropna(subset=["Ma_Hang", "Link_LC"]).copy()
    df["Ma_Hang"] = df["Ma_Hang"].astype(str).str.strip()
    df = df[df["Link_LC"].astype(str).str.startswith("http")].copy()
    df = df.drop_duplicates(subset="Ma_Hang", keep="first")

    if "Mã sp" not in df.columns:
        df["Ma_sp"] = ""
    else:
        df["Ma_sp"] = df["Ma_sp"].fillna("").astype(str).str.strip()

    try:
        df_ia = read_excel_fast(path, sheet_name="import_app")
        df_ia["Mã Hàng"] = df_ia["Mã Hàng"].astype(str).str.strip()
        ia_sub = df_ia[["Mã Hàng", "Tên App", "ID App", "ĐVT App"]].dropna(subset=["Mã Hàng"]).drop_duplicates("Mã Hàng")
        ia_sub = ia_sub.rename(columns={"Mã Hàng": "Ma_Hang", "Tên App": "Ten_App", "ID App": "ID_App", "ĐVT App": "DVT_App"})
        df = df.merge(ia_sub, on="Ma_Hang", how="left")
    except Exception:
        df["Ten_App"] = np.nan
        df["ID_App"] = np.nan
        df["DVT_App"] = np.nan

    if app_file_path and app_file_path.exists():
        try:
            df_app_f = read_excel_fast(app_file_path, sheet_name=0)
            code_c = [c for c in df_app_f.columns if "mã" in str(c).lower()][0]
            id_c = [c for c in df_app_f.columns if "id" in str(c).lower()][0]
            df_app_f[code_c] = df_app_f[code_c].astype(str).str.strip()
            app_dict = df_app_f.set_index(code_c)[id_c].to_dict()
            df["ID_App"] = df["ID_App"].fillna(df["Ma_Hang"].map(app_dict))
        except Exception:
            pass

    df["Ten_App"] = df["Ten_App"].fillna(df["Ten_Hang"])
    df["DVT_App"] = df["DVT_App"].fillna(df["DVT"])

    df["DVT_norm"] = df["DVT"].apply(norm_dvt)
    df["key_exact"] = [norm_key(u, d) for u, d in zip(df["Link_LC"], df["DVT_norm"])]
    df["key_url"] = [norm_key(u) for u in df["Link_LC"]]
    df["Quy_Doi"] = pd.to_numeric(df.get("Quy_Doi", 1.0), errors="coerce").fillna(1.0)
    df["BG_LC_ghi_nhan_truoc"] = pd.to_numeric(df.get("BG_LC_ghi_nhan_truoc", np.nan), errors="coerce")
    return df


def load_he_so(path: Path) -> pd.DataFrame:
    try:
        df = read_excel_fast(path, sheet_name="hệ số  quy cách")
    except Exception:
        df = read_excel_fast(path, sheet_name=0)

    keep = {"Mã Hàng": "Ma_Hang", "* Hệ số": "He_so"}
    for k in keep:
        if k not in df.columns:
            for c in df.columns:
                if "hệ số" in str(c).lower() and k == "* Hệ số":
                    df[k] = df[c]
                elif "mã hàng" in str(c).lower() and k == "Mã Hàng":
                    df[k] = df[c]

    df = df[["Mã Hàng", "* Hệ số"]].rename(columns=keep).dropna(subset=["Ma_Hang"])
    df["Ma_Hang"] = df["Ma_Hang"].astype(str).str.strip()
    df["He_so"] = pd.to_numeric(df["He_so"], errors="coerce").fillna(1.0)
    return df.drop_duplicates(subset="Ma_Hang", keep="first")


def load_category_classification(master_path: Path) -> pd.DataFrame:
    try:
        df_dssp = read_excel_fast(master_path, sheet_name="dssp")
        code_c = [c for c in df_dssp.columns if "mã hàng" in str(c).lower()][0]
        group_c = [c for c in df_dssp.columns if "nhóm hàng" in str(c).lower()][0]
        df_dssp["Ma_Hang"] = df_dssp[code_c].astype(str).str.strip()
        df_dssp["Nhom_hang_master"] = df_dssp[group_c].astype(str).str.strip()
        return df_dssp[["Ma_Hang", "Nhom_hang_master"]].drop_duplicates("Ma_Hang", keep="first")
    except Exception:
        return pd.DataFrame(columns=["Ma_Hang", "Nhom_hang_master"])


def load_crawl(path: Path) -> pd.DataFrame:
    df = read_excel_fast(path, sheet_name=0)
    keep = {"name": "Ten_SP_crawl", "url": "URL_crawl", "packaging": "DVT_crawl", "price": "Gia_crawl"}
    for k in list(keep.keys()):
        if k not in df.columns:
            for c in df.columns:
                if k.lower() in str(c).lower():
                    df[k] = df[c]
                    break

    df = df[list(keep.keys())].rename(columns=keep).copy()
    df = df.dropna(subset=["URL_crawl", "Gia_crawl"]).copy()
    df["Gia_crawl"] = pd.to_numeric(df["Gia_crawl"], errors="coerce")
    df = df[df["Gia_crawl"] > 0].copy()

    df["DVT_crawl_norm"] = df["DVT_crawl"].apply(norm_dvt)
    df["key_exact"] = [norm_key(u, d) for u, d in zip(df["URL_crawl"], df["DVT_crawl_norm"])]
    df["key_url"] = [norm_key(u) for u in df["URL_crawl"]]
    return df


def load_historical_cost_prices_v5(master_path: Path, nhap_path: Path) -> pd.DataFrame:
    """
    COGS Waterfall V5 (TẤT CẢ GIÁ VỐN ĐỀU LÀ ĐƠN GIÁ SAU VAT):
    1. Quét 22 chu kỳ lịch sử trong Sheet 'update giá nhập' (các cột Giá nhập VAT).
    2. Quét file nhập hàng chi tiết PO 1 năm (DanhSachChiTietNhapHang_KV*.xlsx / 28-4.xlsx):
       - Đơn giá vốn sau VAT = Giá nhập + (VAT nhập hàng / Số lượng)
       - Lọc bỏ đơn hủy.
       - Sắp xếp theo Thời gian nhập mới nhất.
       - Ưu tiên phiếu nhập tại 'Kho Tổng CNT' (giá nhập min), sau đó tới chi nhánh khác.
    3. Giá vốn gần nhất = PO mới nhất ưu tiên Kho Tổng (gv_moi_po), nếu không có PO thì lấy gv_moi_sheet.
    4. Giá vốn cũ = Giá vốn gần nhất ngoài file PO (gv_moi_sheet) hoặc gv_cu_sheet.
    """
    # 1. Quét lịch sử 22 kỳ từ Master Sheet (Giá vốn sau VAT)
    df_up = read_excel_fast(master_path, sheet_name="update giá nhập")
    code_c_up = [c for c in df_up.columns if "mã hàng" in str(c).lower()][0]
    df_up["Ma_Hang"] = df_up[code_c_up].astype(str).str.strip()

    vat_cols = [c for c in df_up.columns if any(k in str(c).lower() for k in ["giá nhập vat\n", "giá nhập vat (fix)", "giá nhập vat cũ trước"])]

    def extract_costs_from_history_sheet(row):
        vals = []
        for c in vat_cols:
            v = row[c]
            if pd.notna(v):
                try:
                    num = float(v)
                    if num > 0:
                        vals.append(num)
                except Exception:
                    pass
        if len(vals) >= 2:
            return vals[-2], vals[-1]
        elif len(vals) == 1:
            return vals[0], vals[0]
        return np.nan, np.nan

    res_history = [extract_costs_from_history_sheet(r) for _, r in df_up.iterrows()]
    df_up["gv_cu_sheet"] = [r[0] for r in res_history]
    df_up["gv_moi_sheet"] = [r[1] for r in res_history]
    master_costs = df_up[["Ma_Hang", "gv_cu_sheet", "gv_moi_sheet"]].drop_duplicates("Ma_Hang")

    # 2. Quét file nhập hàng chi tiết PO mới nhất (Tính đơn giá SAU VAT của 1 đơn vị)
    df_nhap = read_excel_fast(nhap_path, sheet_name=0)
    code_col = [c for c in df_nhap.columns if "mã hàng" in str(c).lower()][0]
    df_nhap["Ma_Hang"] = df_nhap[code_col].astype(str).str.strip()

    if "Trạng thái" in df_nhap.columns:
        df_nhap = df_nhap[~df_nhap["Trạng thái"].astype(str).str.contains("Hủy|huỷ", case=False, na=False)].copy()

    time_col = [c for c in df_nhap.columns if "thời gian" in str(c).lower()]
    time_col_name = time_col[0] if time_col else None
    if time_col_name:
        df_nhap["Thời gian dt"] = pd.to_datetime(df_nhap[time_col_name], errors="coerce")
    else:
        df_nhap["Thời gian dt"] = pd.NaT

    # ĐƠN GIÁ VỐN SAU VAT = Giá nhập + (VAT nhập hàng / Số lượng)
    if "Giá nhập" in df_nhap.columns and "VAT nhập hàng" in df_nhap.columns and "Số lượng" in df_nhap.columns:
        qty = pd.to_numeric(df_nhap["Số lượng"], errors="coerce").fillna(1)
        qty = np.where(qty <= 0, 1, qty)
        unit_vat = pd.to_numeric(df_nhap["VAT nhập hàng"], errors="coerce").fillna(0) / qty
        df_nhap["Gia_von_VAT"] = pd.to_numeric(df_nhap["Giá nhập"], errors="coerce").fillna(0) + unit_vat
    elif "Giá nhập" in df_nhap.columns:
        df_nhap["Gia_von_VAT"] = pd.to_numeric(df_nhap["Giá nhập"], errors="coerce")
    elif "Đơn giá" in df_nhap.columns:
        df_nhap["Gia_von_VAT"] = pd.to_numeric(df_nhap["Đơn giá"], errors="coerce")
    else:
        df_nhap["Gia_von_VAT"] = np.nan

    df_nhap = df_nhap[df_nhap["Gia_von_VAT"] > 0].copy()

    # Phân cấp ưu tiên chi nhánh: Kho Tổng = Rank 1, Chi nhánh khác = Rank 2
    branch_col = [c for c in df_nhap.columns if any(k in str(c).lower() for k in ["chi nhánh", "kho", "địa điểm"])]
    if branch_col:
        b_name = branch_col[0]
        df_nhap["is_kho_tong"] = df_nhap[b_name].astype(str).str.contains("Kho Tổng|Kho Tong|KHO TỔNG|KHO TONG", case=False, na=False)
        df_nhap["branch_rank"] = np.where(df_nhap["is_kho_tong"], 1, 2)
    else:
        df_nhap["branch_rank"] = 2
        b_name = None

    dvt_col_po = [c for c in df_nhap.columns if c.lower() in ["đvt", "dvt", "đơn vị tính"]]
    qty_col_po = [c for c in df_nhap.columns if "số lượng" in str(c).lower()]

    sort_cols = []
    if time_col_name:
        sort_cols.append("Thời gian dt")
    sort_cols.extend(["branch_rank", "Gia_von_VAT"])
    ascending_order = [True] * (len(sort_cols) - 2) + [False, True] if len(sort_cols) >= 2 else [True]

    df_nhap_sorted = df_nhap.sort_values(by=sort_cols, ascending=ascending_order)
    po_latest_df = df_nhap_sorted.groupby("Ma_Hang").last().reset_index()

    po_latest = po_latest_df.set_index("Ma_Hang")["Gia_von_VAT"].to_dict()
    po_time_latest = po_latest_df.set_index("Ma_Hang")["Thời gian dt"].dt.strftime("%d/%m/%Y %H:%M:%S").to_dict() if time_col_name else {}
    dvt_po_map = po_latest_df.set_index("Ma_Hang")[dvt_col_po[0]].to_dict() if dvt_col_po else {}
    qty_po_map = po_latest_df.set_index("Ma_Hang")[qty_col_po[0]].to_dict() if qty_col_po else {}
    branch_po_map = po_latest_df.set_index("Ma_Hang")[b_name].to_dict() if b_name else {}

    master_costs["gv_moi_po"] = master_costs["Ma_Hang"].map(po_latest)
    master_costs["Thoi_gian_nhap_gan_nhat"] = master_costs["Ma_Hang"].map(po_time_latest)
    master_costs["DVT_nhap"] = master_costs["Ma_Hang"].map(dvt_po_map)
    master_costs["So_luong_nhap"] = master_costs["Ma_Hang"].map(qty_po_map)
    master_costs["Chi_nhanh_nhap_gan_nhat"] = master_costs["Ma_Hang"].map(branch_po_map)

    master_costs["Gia_von_VAT"] = master_costs["gv_moi_po"].fillna(master_costs["gv_moi_sheet"])
    master_costs["Gia_von_tuan_truoc"] = master_costs["gv_moi_sheet"].where(
        master_costs["gv_moi_po"].notna() & (master_costs["gv_moi_po"] != master_costs["gv_moi_sheet"]),
        master_costs["gv_cu_sheet"],
    )

    return master_costs[["Ma_Hang", "Gia_von_VAT", "Gia_von_tuan_truoc", "Thoi_gian_nhap_gan_nhat", "Chi_nhanh_nhap_gan_nhat", "DVT_nhap", "So_luong_nhap"]]


def load_banggia_medigo(path: Path) -> pd.DataFrame:
    df = read_excel_fast(path, sheet_name=0)
    code_col = [c for c in df.columns if "mã hàng" in str(c).lower()][0]
    df["Ma_Hang"] = df[code_col].astype(str).str.strip()

    price_onl_hcm = pd.to_numeric(df.get("Banggia online HCM", np.nan), errors="coerce")
    price_onl_hn = pd.to_numeric(df.get("Banggia online HN", np.nan), errors="coerce")
    price_chung = pd.to_numeric(df.get("Bảng giá chung", np.nan), errors="coerce")

    df["Gia_ban_hien_tai_Medigo"] = price_onl_hcm.fillna(price_onl_hn).fillna(price_chung)
    df["Ton_kho"] = pd.to_numeric(df.get("Tồn kho", 0), errors="coerce").fillna(0)
    df["Nhom_hang"] = df.get("Nhóm hàng", "").astype(str)

    return df[["Ma_Hang", "Gia_ban_hien_tai_Medigo", "Ton_kho", "Nhom_hang"]].drop_duplicates("Ma_Hang", keep="first")


# ======================================================================
# 4. SMART MULTI-PACKAGING MATCHING
# ======================================================================
def match_competitor_price(mapping: pd.DataFrame, crawl: pd.DataFrame, he_so: pd.DataFrame) -> pd.DataFrame:
    exact_crawl = crawl[["key_exact", "Gia_crawl", "DVT_crawl"]].drop_duplicates("key_exact", keep="first")
    exact_crawl = exact_crawl.rename(columns={"Gia_crawl": "Gia_LC_exact", "DVT_crawl": "DVT_crawl_exact"})
    out = mapping.merge(exact_crawl, on="key_exact", how="left")

    crawl_grouped = crawl.groupby("key_url").agg(
        n_pkgs=("DVT_crawl", "count"),
        dvt_list=("DVT_crawl_norm", list),
        dvt_orig_list=("DVT_crawl", list),
        price_list=("Gia_crawl", list),
    ).reset_index()

    out = out.merge(crawl_grouped, on="key_url", how="left")
    out = out.merge(he_so, on="Ma_Hang", how="left")
    out["He_so"] = out["He_so"].fillna(1.0)

    def resolve_fallback_price(row):
        if pd.notna(row["Gia_LC_exact"]):
            return row["Gia_LC_exact"], row["DVT_crawl_exact"], "khop_chinh_xac"

        if pd.isna(row["n_pkgs"]) or row["n_pkgs"] == 0:
            return np.nan, np.nan, "khong_tim_thay"

        dvt_med = row["DVT_norm"]
        dvt_list = row["dvt_list"]
        dvt_orig_list = row["dvt_orig_list"]
        price_list = row["price_list"]

        if dvt_med in dvt_list:
            idx = dvt_list.index(dvt_med)
            return price_list[idx], dvt_orig_list[idx], "khop_dvt_chuan_hoa"

        if len(price_list) == 1:
            return price_list[0] * row["He_so"], dvt_orig_list[0], "fallback_url_1_dvt"

        hs = row["He_so"]
        if hs != 1.0:
            idx = price_list.index(min(price_list)) if hs > 1 else price_list.index(max(price_list))
            return price_list[idx] * hs, dvt_orig_list[idx], "fallback_he_so_multi_pkg"

        baseline = row["BG_LC_ghi_nhan_truoc"]
        if pd.notna(baseline) and baseline > 0:
            idx = min(range(len(price_list)), key=lambda i: abs(price_list[i] - baseline))
            return price_list[idx], dvt_orig_list[idx], "fallback_closest_baseline"

        max_idx = price_list.index(max(price_list))
        return max(price_list), dvt_orig_list[max_idx], "fallback_default_max"

    resolved = [resolve_fallback_price(r) for _, r in out.iterrows()]
    out["BG_LC_hien_tai"] = [r[0] for r in resolved]
    out["DVT_crawl_matched"] = [r[1] for r in resolved]
    out["Nguon_khop_gia"] = [r[2] for r in resolved]

    return out[[
        "Ma_Hang", "Ma_sp", "Ten_Hang", "DVT", "ID_App", "Ten_App", "DVT_App",
        "Link_LC", "BG_LC_hien_tai", "DVT_crawl_matched", "He_so", "Nguon_khop_gia", "BG_LC_ghi_nhan_truoc", "Quy_Doi", "DVT_norm"
    ]]


# ======================================================================
# 5. AUTOMATED PRICING ENGINE V5 (2 TẦNG ĐỊNH GIÁ)
# ======================================================================
def calculate_pricing_v5(row) -> tuple:
    gv = row["Gia_von_VAT"]
    bg_lc = row["BG_LC_hien_tai"]
    ma_sp = str(row.get("Ma_sp", "")).upper().strip()
    nhom = str(row.get("Nhom_hang", "")).upper().strip()

    if pd.isna(gv) or gv <= 0:
        p_onl = bg_lc if pd.notna(bg_lc) and bg_lc > 0 else np.nan
        return p_onl, np.nan, np.nan, 0, np.nan, np.nan, "Chưa có giá vốn"

    # 1. Xác định Target Margin theo Bảng C-F
    if "CK" in ma_sp or "CHIẾT KHẤU" in nhom:
        m_onl, m_low, m_high = 0.58, 0.58, 0.58
    elif "MM" in ma_sp:
        m_onl, m_low, m_high = 0.25, 0.25, 0.25
    elif "LL" in ma_sp:
        m_onl, m_low, m_high = 0.14, 0.20, 0.20
    elif "TPCN" in nhom or "THỰC PHẨM" in nhom:
        m_onl, m_low, m_high = 0.15, 0.18, 0.18
    elif "TBYT" in nhom or "THIẾT BỊ" in nhom or "FMCG" in nhom or "TIÊU DÙNG" in nhom:
        m_onl, m_low, m_high = 0.20, 0.28, 0.28
    else:  # Mặc định Thuốc (NY1, NY2, NY3)
        m_onl, m_low, m_high = 0.14, 0.16, 0.20

    # Giá theo Bảng C-F (làm tròn lên hàng nghìn)
    p_cf_onl = np.ceil((gv / (1 - m_onl)) / 1000) * 1000 if m_onl < 1 else gv
    p_cf_low = np.ceil((gv / (1 - m_low)) / 1000) * 1000 if m_low < 1 else gv
    p_cf_high = np.ceil((gv / (1 - m_high)) / 1000) * 1000 if m_high < 1 else gv

    # 2. Định giá Tầng 1 (Khi CÓ giá đối thủ)
    if pd.notna(bg_lc) and bg_lc > 0:
        if bg_lc <= 100000:
            pct_over = 0.10
        elif bg_lc <= 300000:
            pct_over = 0.08
        elif bg_lc <= 500000:
            pct_over = 0.06
        elif bg_lc <= 1000000:
            pct_over = 0.04
        else:
            pct_over = 0.025

        margin_at_lc = (bg_lc - gv) / bg_lc

        # Nếu Margin tại giá LC >= Target Margin ngành -> Bán đúng bằng đối thủ
        if margin_at_lc >= m_onl:
            p_onl = bg_lc
            reason = f"Bán = LC (Margin {margin_at_lc:.1%} >= Target {m_onl:.0%})"
        else:
            # Margin mỏng hoặc Giá vốn >= Đối thủ -> Áp dụng Range đối thủ
            p_cand = bg_lc * (1 + pct_over)
            p_round = np.ceil(p_cand / 1000) * 1000

            # Chống bán lỗ: Nếu sau khi cộng range vẫn < giá vốn -> Fallback Bảng C-F
            if p_round < gv:
                p_onl = p_cf_onl
                reason = "Fallback Bảng C-F (Chống bán lỗ)"
            else:
                p_onl = p_round
                reason = f"Cộng Range +{pct_over*100:.1f}% (Margin tại LC chỉ {margin_at_lc:.1%})"
    else:
        # Tầng 2: Không có giá đối thủ -> Bảng C-F
        p_onl = p_cf_onl
        reason = "Định giá theo Bảng C-F (Không có giá LC)"

    so_voi_dt = (p_onl - bg_lc) if pd.notna(bg_lc) and bg_lc > 0 else np.nan
    margin_calc = (p_onl - gv) / p_onl if p_onl > 0 else 0
    loi_nhuan_gop = p_onl - gv if p_onl > 0 else 0

    return p_onl, p_cf_low, p_cf_high, so_voi_dt, margin_calc, loi_nhuan_gop, reason


# ======================================================================
# 6. PIPELINE CHẠY CHÍNH V5
# ======================================================================
def run_pricing_pipeline():
    print("=" * 105)
    print("      MEDIGO PRICING ENGINE V5 (COGS WATERFALL KHO TỔNG & 2-TIER AUTOMATED PRICING)")
    print("=" * 105)

    print("\n--- [1/6] Quét các file dữ liệu mới nhất trong DATA/ ---")
    crawl_file = find_latest_file(DATA_DIR, ["pharmacity_products.xlsx", "*crawl*.xlsx", "*pharmacity*.xlsx"], exclude_keywords=["mapping", "he_so", "banggia"])
    nhap_file = find_latest_file(DATA_DIR, ["*DanhSachChiTietNhapHang*.xlsx", "28-4.xlsx", "*nhaphang*.xlsx"], exclude_keywords=["mapping", "banggia_med"])
    mapping_file = find_latest_file(DATA_DIR, ["banggia_med_data_nhaphang_final.xlsx"])
    he_so_file = find_latest_file(DATA_DIR, ["crawl - mapping LC.xlsx", "crawl_-_mapping_LC.xlsx", "*he_so*.xlsx"])
    kiot_price_file = find_latest_file(DATA_DIR, ["BangGia_KV*.xlsx", "*BangGia*.xlsx"])

    app_mapping_file = DATA_DIR / "mapping Kiot - MedigoApp.xlsx"
    app_file_path = app_mapping_file if app_mapping_file.exists() else None

    print(f"  • File Crawl Đối thủ   : {crawl_file.name}")
    print(f"  • File Nhập hàng PO    : {nhap_file.name}")
    print(f"  • File Bảng giá Kiot   : {kiot_price_file.name}")
    print(f"  • File Master Mapping  : {mapping_file.name}")
    print(f"  • File Hệ số Quy cách  : {he_so_file.name}")

    print("\n--- [2/6] Nạp Master Data, Lịch sử Giá Vốn V5 (Ưu tiên Kho Tổng) & Smart Matching ---")
    mapping = load_lc_mapping(mapping_file, app_file_path)
    he_so = load_he_so(he_so_file)
    categories = load_category_classification(mapping_file)
    crawl = load_crawl(crawl_file)
    lc_priced = match_competitor_price(mapping, crawl, he_so)

    cost_history = load_historical_cost_prices_v5(mapping_file, nhap_file)
    banggia_kv = load_banggia_medigo(kiot_price_file)

    snap = lc_priced.merge(cost_history, on="Ma_Hang", how="left")
    snap = snap.merge(banggia_kv, on="Ma_Hang", how="left")
    snap = snap.merge(categories, on="Ma_Hang", how="left")

    snap["Nhom_hang"] = snap["Nhom_hang"].where(snap["Nhom_hang"] != "", snap["Nhom_hang_master"])

    print("\n--- [3/6] So sánh biến động với Tuần trước & Phân loại ---")
    first_run = not STATE_FILE.exists()
    if first_run:
        print("=> Lần đầu chạy V5: Sử dụng 'BG LC' ghi nhận trước làm baseline đối thủ.")
        snap_prev = pd.DataFrame(columns=["Ma_Hang", "BG_LC_hien_tai"])
    else:
        backup_name = HISTORY_DIR / f"state_backup_v5_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        shutil.copy2(STATE_FILE, backup_name)
        snap_prev = pd.read_csv(STATE_FILE)
        snap_prev["Ma_Hang"] = snap_prev["Ma_Hang"].astype(str).str.strip()

    cmp_df = snap.merge(
        snap_prev[["Ma_Hang", "BG_LC_hien_tai"]],
        on="Ma_Hang", how="left", suffixes=("", "_tuan_truoc"),
    )
    cmp_df = cmp_df.rename(columns={
        "BG_LC_hien_tai_tuan_truoc": "BG_LC_tuan_truoc",
        "Gia_ban_hien_tai_Medigo": "Gia_ban_truoc_do",
    })

    dung_baseline_mapping = cmp_df["BG_LC_tuan_truoc"].isna() & cmp_df["BG_LC_ghi_nhan_truoc"].notna()
    cmp_df.loc[dung_baseline_mapping, "BG_LC_tuan_truoc"] = cmp_df.loc[dung_baseline_mapping, "BG_LC_ghi_nhan_truoc"]

    def khac_nhau(a, b, tol=0.01):
        if pd.isna(a) and pd.isna(b):
            return False
        if pd.isna(a) or pd.isna(b):
            return True
        return abs(a - b) > tol

    khac_lc = np.array([khac_nhau(a, b) for a, b in zip(cmp_df["BG_LC_hien_tai"], cmp_df["BG_LC_tuan_truoc"])])
    khac_gv = np.array([khac_nhau(a, b) for a, b in zip(cmp_df["Gia_von_VAT"], cmp_df["Gia_von_tuan_truoc"])])

    cmp_df["Doi_thu_doi_gia"] = cmp_df["BG_LC_hien_tai"].notna() & cmp_df["BG_LC_tuan_truoc"].notna() & khac_lc
    cmp_df["Mat_gia_tuan_nay"] = cmp_df["BG_LC_hien_tai"].isna() & cmp_df["BG_LC_tuan_truoc"].notna()
    cmp_df["Gia_von_doi"] = cmp_df["Gia_von_VAT"].notna() & cmp_df["Gia_von_tuan_truoc"].notna() & khac_gv

    ROUND_FACTORS = [2, 3, 4, 5, 6, 8, 10, 12, 15, 20, 24, 30, 50, 60, 100]

    def nghi_lech_dvt(gia_moi, gia_cu, quy_doi):
        if pd.isna(gia_moi) or pd.isna(gia_cu) or gia_moi == 0 or gia_cu == 0:
            return False
        ty_le = max(gia_moi, gia_cu) / min(gia_moi, gia_cu)
        candidates = ROUND_FACTORS + ([quy_doi] if pd.notna(quy_doi) and quy_doi > 1 else [])
        return any(abs(ty_le - c) / c < 0.05 for c in candidates)

    cmp_df["Nghi_lech_DVT"] = [
        nghi_lech_dvt(a, b, q)
        for a, b, q in zip(cmp_df["BG_LC_hien_tai"], cmp_df["BG_LC_tuan_truoc"], cmp_df["Quy_Doi"])
    ]
    cmp_df["Doi_thu_doi_gia_that"] = cmp_df["Doi_thu_doi_gia"] & ~cmp_df["Nghi_lech_DVT"]
    cmp_df["Nghi_lech_DVT"] = cmp_df["Doi_thu_doi_gia"] & cmp_df["Nghi_lech_DVT"]

    def gen_reason(r):
        reasons = []
        if r["Doi_thu_doi_gia_that"]:
            bg_moi = r["BG_LC_hien_tai"]
            bg_cu = r["BG_LC_tuan_truoc"]
            pct = (bg_moi - bg_cu) / bg_cu if bg_cu > 0 else 0
            act = "tăng" if pct > 0 else "giảm"
            reasons.append(f"Đối thủ {act} giá ({pct:+.1%})")
        if r["Gia_von_doi"]:
            gv_moi = r["Gia_von_VAT"]
            gv_cu = r["Gia_von_tuan_truoc"]
            pct_gv = (gv_moi - gv_cu) / gv_cu if gv_cu > 0 else 0
            act_gv = "tăng" if pct_gv > 0 else "giảm"
            reasons.append(f"Giá vốn {act_gv} ({pct_gv:+.1%})")
        return " & ".join(reasons) if reasons else ""

    cmp_df["Ly_do_dieu_chinh"] = cmp_df.apply(gen_reason, axis=1)

    print("\n--- [4/6] Chạy Ma trận Định giá Tự động V5 (Online, Offline Low, Offline High) ---")
    pricing_results = [calculate_pricing_v5(r) for _, r in cmp_df.iterrows()]
    cmp_df["Gia_moi_Online"] = [r[0] for r in pricing_results]
    cmp_df["Gia_moi_offline_low"] = [r[1] for r in pricing_results]
    cmp_df["Gia_moi_offline_hight"] = [r[2] for r in pricing_results]
    cmp_df["So_voi_doi_thu"] = [r[3] for r in pricing_results]
    cmp_df["Margin"] = [r[4] for r in pricing_results]
    cmp_df["Loi_nhuan_gop_VND"] = [r[5] for r in pricing_results]
    cmp_df["Chi_tiet_dinh_gia"] = [r[6] for r in pricing_results]

    # Cảnh báo hàng nhập giá cao: Giá vốn Medigo > Giá bán Đối thủ
    cmp_df["Nhap_gia_cao"] = (
        cmp_df["Gia_von_VAT"].notna()
        & cmp_df["BG_LC_hien_tai"].notna()
        & (cmp_df["Gia_von_VAT"] > cmp_df["BG_LC_hien_tai"])
    )

    # Lọc bỏ sản phẩm ngừng kinh doanh: Tồn kho <= 0 & Không có PO & Có giá bán
    is_removed = (
        (cmp_df["Ton_kho"] <= 0)
        & (cmp_df["Thoi_gian_nhap_gan_nhat"].isna())
        & (cmp_df["Gia_ban_truoc_do"] > 0)
    )

    df_active = cmp_df[~is_removed].copy()
    df_removed = cmp_df[is_removed].copy()

    print(f"  • Tổng số sản phẩm Master ban đầu : {len(cmp_df):,} sản phẩm")
    print(f"  • Số sản phẩm bị loại (Tồn=0, ko PO): {len(df_removed):,} sản phẩm")
    print(f"  • Số sản phẩm Active đưa vào định giá: {len(df_active):,} sản phẩm")

    # Tách các bảng lọc theo điều kiện trên tập Active
    df_doi_thu_doi = df_active[df_active["Doi_thu_doi_gia_that"]].copy()
    df_gia_von_doi = df_active[df_active["Gia_von_doi"]].copy()
    df_doi_ca_2 = df_active[df_active["Doi_thu_doi_gia_that"] & df_active["Gia_von_doi"]].copy()
    df_tong_hop = df_active[df_active["Doi_thu_doi_gia_that"] | df_active["Gia_von_doi"]].copy()
    df_nhap_gia_cao = df_active[df_active["Nhap_gia_cao"]].copy()

    # Lưu snapshot tuần này
    df_active[["Ma_Hang", "BG_LC_hien_tai", "Gia_von_VAT"]].to_csv(STATE_FILE, index=False)

    # ======================================================================
    # 7. XUẤT 2 BẢN BÁO CÁO EXCEL V5 (ONLINE-ONLY & FULL)
    # ======================================================================
    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M")

    # 1. BẢN ONLINE-ONLY (Để trống các cột Offline Low, Offline High theo yêu cầu)
    output_filename_online = OUTPUT_DIR / f"Bao_cao_import_price_V5_OnlineOnly_{timestamp_str}.xlsx"
    # 2. BẢN FULL (Có gợi ý đầy đủ cả Online, Offline Low, Offline High)
    output_filename_full = OUTPUT_DIR / f"Bao_cao_import_price_V5_Full_{timestamp_str}.xlsx"

    cols_export_template = [
        "Ma_Hang", "Ma_sp", "Ten_Hang", "DVT", "DVT_nhap", "So_luong_nhap", "Chi_nhanh_nhap_gan_nhat",
        "ID_App", "Ten_App", "DVT_App",
        "BG_LC_tuan_truoc", "Link_LC", "BG_LC_hien_tai", "DVT_crawl_matched", "He_so",
        "Gia_von_VAT", "Gia_von_tuan_truoc", "Thoi_gian_nhap_gan_nhat", "Gia_ban_truoc_do",
        "Gia_moi_Online", "Gia_moi_offline_low", "Gia_moi_offline_hight",
        "So_voi_doi_thu", "Margin", "Loi_nhuan_gop_VND", "Ly_do_dieu_chinh"
    ]

    rename_export = {
        "Ma_Hang": "Mã Hàng (Kiot)",
        "Ma_sp": "Mã sp",
        "Ten_Hang": "Tên Hàng (Kiot)",
        "DVT": "ĐVT (Kiot)",
        "DVT_nhap": "ĐVT nhập (PO)",
        "So_luong_nhap": "Số lượng nhập",
        "Chi_nhanh_nhap_gan_nhat": "Chi nhánh nhập gần nhất",
        "ID_App": "ID App",
        "Ten_App": "Tên trên App",
        "DVT_App": "ĐVT trên App",
        "BG_LC_tuan_truoc": "BG LC",
        "Link_LC": "Link LC",
        "BG_LC_hien_tai": "BG LC hiện tại",
        "DVT_crawl_matched": "ĐVT Đối thủ (Crawl)",
        "He_so": "Hệ số quy cách",
        "Gia_von_VAT": "Giá vốn gần nhất",
        "Gia_von_tuan_truoc": "Giá vốn cũ",
        "Thoi_gian_nhap_gan_nhat": "Thời gian nhập gần nhất",
        "Gia_ban_truoc_do": "Gía bán trước đó",
        "Gia_moi_Online": "Giá mới Online",
        "Gia_moi_offline_low": "Giá mới offline low",
        "Gia_moi_offline_hight": "Giá mới offline hight",
        "So_voi_doi_thu": "So với đối thủ",
        "Margin": "Margin",
        "Loi_nhuan_gop_VND": "Lợi nhuận gộp (VNĐ)",
        "Ly_do_dieu_chinh": "Lý do điều chỉnh",
    }

    def export_workbook(target_path: Path, df_source: pd.DataFrame, df_rem: pd.DataFrame, is_online_only: bool = False):
        df_act = df_source.copy()
        df_r = df_rem.copy()
        if is_online_only:
            df_act["Gia_moi_offline_low"] = np.nan
            df_act["Gia_moi_offline_hight"] = np.nan
            df_r["Gia_moi_offline_low"] = np.nan
            df_r["Gia_moi_offline_hight"] = np.nan

        df_dt = df_act[df_act["Doi_thu_doi_gia_that"]].copy()
        df_gv = df_act[df_act["Gia_von_doi"]].copy()
        df_both = df_act[df_act["Doi_thu_doi_gia_that"] & df_act["Gia_von_doi"]].copy()
        df_all_v = df_act[df_act["Doi_thu_doi_gia_that"] | df_act["Gia_von_doi"]].copy()
        df_high_cost = df_act[df_act["Nhap_gia_cao"]].copy()

        with pd.ExcelWriter(target_path, engine="openpyxl") as writer:
            # Sheet 1: Đối thủ đổi giá (Trình Sếp)
            df_dt[cols_export_template].rename(columns=rename_export).to_excel(
                writer, sheet_name="Doi thu doi gia (Trinh Sep)", index=False
            )
            # Sheet 2: Giá vốn thay đổi (Trình Sếp)
            df_gv[cols_export_template].rename(columns=rename_export).to_excel(
                writer, sheet_name="Gia von doi (Trinh Sep)", index=False
            )
            # Sheet 3: Vừa đổi giá đối thủ vừa đổi giá vốn
            df_both[cols_export_template].rename(columns=rename_export).to_excel(
                writer, sheet_name="Doi ca 2 gia (Doi thu & GV)", index=False
            )
            # Sheet 4: Tổng hợp toàn bộ biến động
            df_all_v[cols_export_template].rename(columns=rename_export).to_excel(
                writer, sheet_name="Tong hop bien dong", index=False
            )
            # Sheet 5: Danh sách nhập giá cao (Giá vốn Medigo > Giá bán Đối thủ)
            df_high_cost[cols_export_template].rename(columns=rename_export).to_excel(
                writer, sheet_name="List sp gia nhap cao", index=False
            )
            # Sheet 6: Bảng so sánh 4 giá & Đề xuất định giá
            df_act[[
                "Ma_Hang", "Ten_Hang", "DVT", "DVT_nhap", "So_luong_nhap", "Chi_nhanh_nhap_gan_nhat",
                "ID_App", "Ten_App", "DVT_App",
                "BG_LC_tuan_truoc", "BG_LC_hien_tai", "Gia_von_tuan_truoc", "Gia_von_VAT", "Thoi_gian_nhap_gan_nhat", "Gia_ban_truoc_do",
                "Gia_moi_Online", "Gia_moi_offline_low", "Gia_moi_offline_hight", "So_voi_doi_thu", "Margin", "Loi_nhuan_gop_VND"
            ]].rename(columns={
                "Ma_Hang": "Mã Hàng (Kiot)", "Ten_Hang": "Tên Hàng (Kiot)", "DVT": "ĐVT (Kiot)",
                "DVT_nhap": "ĐVT nhập (PO)", "So_luong_nhap": "Số lượng nhập", "Chi_nhanh_nhap_gan_nhat": "Chi nhánh nhập gần nhất",
                "ID_App": "ID App", "Ten_App": "Tên trên App", "DVT_App": "ĐVT trên App",
                "BG_LC_tuan_truoc": "Giá đối thủ TUẦN TRƯỚC",
                "BG_LC_hien_tai": "Giá đối thủ HIỆN TẠI",
                "Gia_von_tuan_truoc": "Giá vốn có VAT TRƯỚC",
                "Gia_von_VAT": "Giá vốn có VAT SAU",
                "Thoi_gian_nhap_gan_nhat": "Thời gian nhập gần nhất",
                "Gia_ban_truoc_do": "Giá bán hiện tại Medigo",
                "Gia_moi_Online": "Giá mới Online",
                "Gia_moi_offline_low": "Giá mới offline low",
                "Gia_moi_offline_hight": "Giá mới offline hight",
                "So_voi_doi_thu": "So với đối thủ",
                "Margin": "Margin",
                "Loi_nhuan_gop_VND": "Lợi nhuận gộp (VNĐ)",
            }).to_excel(writer, sheet_name="Bang so sanh 4 gia & De xuat", index=False)
            # Sheet 7: Danh sách bị loại (Đối soát QA)
            df_r[cols_export_template].rename(columns=rename_export).to_excel(
                writer, sheet_name="DS bi loai (Ton=0, ko PO)", index=False
            )
            # Sheet 8: Nghi lệch ĐVT (QA)
            df_act[df_act["Nghi_lech_DVT"]][cols_export_template].rename(columns=rename_export).to_excel(
                writer, sheet_name="Nghi lech DVT (QA)", index=False
            )
            # Sheet 9: Mất giá tuần này (QA)
            df_act[df_act["Mat_gia_tuan_nay"]][cols_export_template].rename(columns=rename_export).to_excel(
                writer, sheet_name="Mat gia tuan nay (QA)", index=False
            )
            # Sheet 10: Toàn bộ danh mục Master đã lọc (Active)
            df_act[cols_export_template].rename(columns=rename_export).to_excel(
                writer, sheet_name="Toan bo tuan nay (V5)", index=False
            )

    print(f"\n--- [5/6] Xuất 2 file Báo cáo V5 (10 Sheet) ---")
    print(f"  • Xuất Bản 1 (Không điền giá Offline): {output_filename_online.name}")
    export_workbook(output_filename_online, df_active, df_removed, is_online_only=True)

    print(f"  • Xuất Bản 2 (Đầy đủ giá gợi ý):       {output_filename_full.name}")
    export_workbook(output_filename_full, df_active, df_removed, is_online_only=False)

    print("\n--- [6/6] Hoàn tất & Hiển thị 10 Mã tiêu biểu Đối thủ đổi giá ---")
    print("=" * 135)
    print(f"{'Mã Kiot':<11} | {'ĐVT':<5} | {'BG LC Cũ':<9} | {'BG LC Mới':<9} | {'GV Cũ':<9} | {'GV Mới':<9} | {'Giá Mới Onl':<11} | {'Margin':<7} | {'Lý do điều chỉnh'}")
    print("-" * 135)
    for _, r in df_doi_thu_doi.head(10).iterrows():
        bg_cu = f"{r['BG_LC_tuan_truoc']:,.0f}" if pd.notna(r['BG_LC_tuan_truoc']) else "-"
        bg_moi = f"{r['BG_LC_hien_tai']:,.0f}" if pd.notna(r['BG_LC_hien_tai']) else "-"
        gv_cu = f"{r['Gia_von_tuan_truoc']:,.0f}" if pd.notna(r['Gia_von_tuan_truoc']) else "-"
        gv_moi = f"{r['Gia_von_VAT']:,.0f}" if pd.notna(r['Gia_von_VAT']) else "-"
        p_onl_str = f"{r['Gia_moi_Online']:,.0f}" if pd.notna(r['Gia_moi_Online']) else "-"
        margin_str = f"{r['Margin']:.1%}" if pd.notna(r['Margin']) else "-"
        print(f"{r['Ma_Hang']:<11} | {r['DVT']:<5} | {bg_cu:<9} | {bg_moi:<9} | {gv_cu:<9} | {gv_moi:<9} | {p_onl_str:<11} | {margin_str:<7} | {r['Ly_do_dieu_chinh']}")

    print("=" * 135)
    print(f"• Bản 1 (Không điền giá Offline - Dành cho Sếp duyệt): {output_filename_online}")
    print(f"• Bản 2 (Đầy đủ cả giá Online + Offline gợi ý):        {output_filename_full}")
    print(f"• Tổng danh mục Active đưa vào định giá:               {len(df_active):,} mã")
    print(f"• Số mã Đối thủ thay đổi giá THẬT (Sheet 1):          {len(df_doi_thu_doi):,} mã")
    print(f"• Số mã Giá vốn nhập hàng thay đổi (Sheet 2):          {len(df_gia_von_doi):,} mã")
    print(f"• Số mã Vừa đổi giá đối thủ vừa đổi giá vốn (Sheet 3): {len(df_doi_ca_2):,} mã")
    print(f"• Số mã Giá vốn Medigo > Giá bán Đối thủ (Sheet 5):    {len(df_nhap_gia_cao):,} mã")
    print(f"• Danh mục Đã loại bỏ (Tồn=0, ko PO) (Sheet 7):        {len(df_removed):,} mã")
    print("=" * 135)

    return output_filename_online


if __name__ == "__main__":
    run_pricing_pipeline()
