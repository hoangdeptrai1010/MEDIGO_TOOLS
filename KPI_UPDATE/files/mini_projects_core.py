# -*- coding: utf-8 -*-
"""
mini_projects_core.py
======================
Engine DÙNG CHUNG cho mọi "dự án mini" (KAT, Ladycare, Party Smart, AVC,
và bất kỳ dự án nào công ty thêm vào các tháng sau — 3 hay 7 dự án đều chạy
được không cần sửa code).

Ý TƯỞNG CỐT LÕI
---------------
Dù tên dự án là gì, cơ chế thưởng luôn rơi vào 1 khuôn chung gồm 2 phần:

  1. "gate_on"  : điều kiện ĐẠT được tính theo gì — 'qty' (số lượng) hay
                  'revenue' (doanh thu) của đúng nhóm SKU dự án đó.
  2. "reward"   : phần thưởng khi đạt — có 3 kiểu:
        - "per_unit" : thưởng = rate * giá_trị(basis)   (basis: qty hoặc revenue)
        - "percent"  : thưởng = rate * giá_trị(basis)   (basis luôn là revenue, rate là %)
        - "flat"     : thưởng = số tiền cố định

  Ví dụ 4 dự án thật (tháng 8/2026) đều diễn tả được bằng khuôn này:
    - KAT         : gate_on=qty,      reward=per_unit (basis=qty)
    - Ladycare    : gate_on=revenue,  reward=percent  (basis=revenue)
    - Party Smart : gate_on=qty,      reward=per_unit (basis=qty)
    - AVC         : gate_on=revenue,  reward=per_unit (basis=qty)  <- ngưỡng theo doanh thu
                                                                       nhưng trả thưởng theo SL

Vì tất cả quy về 1 khuôn, engine chỉ cần viết 1 LẦN DUY NHẤT, không cần biết
trước tên/số lượng dự án. Thêm dự án thứ 5, 6, 7 chỉ cần thêm 1 object JSON
đúng khuôn vào config, KHÔNG đụng vào file .py này.

CÁCH DÙNG
---------
    from mini_projects_core import MiniProjectsEngine

    engine = MiniProjectsEngine(cfg["mini_projects"])          # đọc từ monthly_config.json
    engine.scan_invoice(inv_rows, header_map)                  # quét hóa đơn 1 lần
    engine.apply_returns(return_rows, header_map)               # trừ trả hàng (nếu có)

    for seller, branch in engine.all_sellers():
        for proj in engine.projects:
            qty, rev = engine.get_seller_value(seller, proj.id)
            bonus    = engine.calc_individual_bonus(proj.id, seller)
            ...
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple


# --------------------------------------------------------------------------- #
# 1. Cấu trúc dữ liệu 1 "tier" (1 mốc thưởng) — dùng chung cho cá nhân & cửa hàng
# --------------------------------------------------------------------------- #

@dataclass
class Tier:
    gate_on: str            # 'qty' hoặc 'revenue' — điều kiện để ĐẠT mốc này
    min_value: float        # ngưỡng tối thiểu để đạt mốc
    reward_type: str        # 'per_unit' | 'percent' | 'flat'
    reward_basis: str       # 'qty' | 'revenue' (bỏ qua nếu reward_type == 'flat')
    rate: float             # đơn giá / tỷ lệ % / số tiền cố định (tuỳ reward_type)
    min_distinct_sku: int = 0   # điều kiện phụ (VD: KAT yêu cầu tối thiểu 3 mã khác nhau)
    label: str = ""

    @staticmethod
    def from_dict(d: dict) -> "Tier":
        return Tier(
            gate_on=d.get("gate_on", "qty"),
            min_value=float(d.get("min", 0)),
            reward_type=d.get("reward", {}).get("type", "per_unit"),
            reward_basis=d.get("reward", {}).get("basis", "qty"),
            rate=float(d.get("reward", {}).get("rate", 0)),
            min_distinct_sku=int(d.get("min_distinct_sku", 0)),
            label=d.get("label", ""),
        )

    def is_met(self, qty: float, revenue: float, distinct_sku: int = 0) -> bool:
        val = qty if self.gate_on == "qty" else revenue
        if val < self.min_value:
            return False
        if self.min_distinct_sku and distinct_sku < self.min_distinct_sku:
            return False
        return True

    def compute_reward(self, qty: float, revenue: float) -> float:
        if self.reward_type == "flat":
            return self.rate
        basis_val = qty if self.reward_basis == "qty" else revenue
        return basis_val * self.rate


def best_tier_reward(tiers: List[Tier], qty: float, revenue: float,
                      distinct_sku: int = 0) -> float:
    """
    Trong 1 danh sách tier (đã sắp theo min_value tăng dần hay không đều được),
    tìm tier CAO NHẤT mà điều kiện đạt được, trả về số tiền thưởng.
    Đây là hàm DUY NHẤT xử lý logic tier cho MỌI dự án — không hardcode tên dự án.
    """
    best_reward = 0.0
    best_threshold = -1.0
    for t in tiers:
        if t.is_met(qty, revenue, distinct_sku) and t.min_value >= best_threshold:
            best_threshold = t.min_value
            best_reward = t.compute_reward(qty, revenue)
    return best_reward


# --------------------------------------------------------------------------- #
# 2. Định nghĩa 1 dự án mini (cho 1 vùng HN hoặc HCM)
# --------------------------------------------------------------------------- #

@dataclass
class MiniProjectRegion:
    skus: List[str]
    individual_tiers: List[Tier]
    store_tiers: List[Tier]


@dataclass
class MiniProject:
    id: str
    name: str
    regions: Dict[str, MiniProjectRegion]   # key: 'HN' | 'HCM'

    @staticmethod
    def from_dict(d: dict) -> "MiniProject":
        regions = {}
        for region_key, region_def in d.get("regions", {}).items():
            regions[region_key] = MiniProjectRegion(
                skus=list(region_def.get("skus", [])),
                individual_tiers=[Tier.from_dict(t) for t in region_def.get("individual_tiers", [])],
                store_tiers=[Tier.from_dict(t) for t in region_def.get("store_tiers", [])],
            )
        return MiniProject(id=d["id"], name=d.get("name", d["id"]), regions=regions)


# --------------------------------------------------------------------------- #
# 3. Engine chính — quét hóa đơn, cộng dồn, tính thưởng
# --------------------------------------------------------------------------- #

class MiniProjectsEngine:
    def __init__(self, mini_projects_cfg: List[dict], hn_branches=("Hàng Bông", "Đường Láng")):
        self.projects: List[MiniProject] = [MiniProject.from_dict(p) for p in mini_projects_cfg]
        self.hn_branches = set(hn_branches)

        # sku -> (project_id) ; kiểm tra trùng SKU giữa các dự án ngay khi khởi tạo
        self.sku_to_project: Dict[str, str] = {}
        dup_warnings = []
        for proj in self.projects:
            for region_def in proj.regions.values():
                for sku in region_def.skus:
                    if sku in self.sku_to_project and self.sku_to_project[sku] != proj.id:
                        dup_warnings.append(
                            f"SKU '{sku}' vừa thuộc dự án '{self.sku_to_project[sku]}' "
                            f"vừa thuộc dự án '{proj.id}' -> KIỂM TRA LẠI CONFIG."
                        )
                    self.sku_to_project[sku] = proj.id
        self.dup_warnings = dup_warnings
        for w in dup_warnings:
            print(f"[MiniProjectsEngine][CẢNH BÁO] {w}")

        # (seller, project_id) -> {'qty':..,'rev':..,'branch':..,'skus':set()}
        self._seller_data: Dict[Tuple[str, str], dict] = {}
        # (branch, project_id) -> {'qty':..,'rev':..,'skus':set()}
        self._branch_data: Dict[Tuple[str, str], dict] = {}

    def region_of(self, branch: str) -> str:
        return "HN" if branch in self.hn_branches else "HCM"

    def project_of_sku(self, sku: str) -> Optional[str]:
        return self.sku_to_project.get(sku)

    # ---- Quét dữ liệu ----------------------------------------------------- #

    def add_sale(self, seller: str, branch: str, sku: str, qty: float, revenue: float):
        """Gọi 1 lần cho mỗi dòng hóa đơn (đã lọc theo sku thuộc dự án nào)."""
        proj_id = self.project_of_sku(sku)
        if proj_id is None or not seller:
            return
        sk = (seller, proj_id)
        if sk not in self._seller_data:
            self._seller_data[sk] = {"qty": 0.0, "rev": 0.0, "branch": branch, "skus": set()}
        self._seller_data[sk]["qty"] += qty
        self._seller_data[sk]["rev"] += revenue
        self._seller_data[sk]["skus"].add(sku)

        bk = (branch, proj_id)
        if bk not in self._branch_data:
            self._branch_data[bk] = {"qty": 0.0, "rev": 0.0, "skus": set()}
        self._branch_data[bk]["qty"] += qty
        self._branch_data[bk]["rev"] += revenue
        self._branch_data[bk]["skus"].add(sku)

    def subtract_return(self, seller: str, branch: str, sku: str, qty: float, revenue: float):
        """Trừ hàng trả lại — không cho âm."""
        proj_id = self.project_of_sku(sku)
        if proj_id is None or not seller:
            return
        sk = (seller, proj_id)
        if sk in self._seller_data:
            self._seller_data[sk]["qty"] = max(0.0, self._seller_data[sk]["qty"] - qty)
            self._seller_data[sk]["rev"] = max(0.0, self._seller_data[sk]["rev"] - revenue)
        bk = (branch, proj_id)
        if bk in self._branch_data:
            self._branch_data[bk]["qty"] = max(0.0, self._branch_data[bk]["qty"] - qty)
            self._branch_data[bk]["rev"] = max(0.0, self._branch_data[bk]["rev"] - revenue)

    # ---- Truy vấn kết quả --------------------------------------------------- #

    def get_seller_value(self, seller: str, project_id: str) -> Tuple[float, float, int]:
        d = self._seller_data.get((seller, project_id))
        if not d:
            return 0.0, 0.0, 0
        return d["qty"], d["rev"], len(d["skus"])

    def get_branch_value(self, branch: str, project_id: str) -> Tuple[float, float, int]:
        d = self._branch_data.get((branch, project_id))
        if not d:
            return 0.0, 0.0, 0
        return d["qty"], d["rev"], len(d["skus"])

    def calc_individual_bonus(self, project_id: str, seller: str, branch: str) -> float:
        proj = self._get_project(project_id)
        region = proj.regions.get(self.region_of(branch))
        if not region:
            return 0.0
        qty, rev, distinct = self.get_seller_value(seller, project_id)
        return best_tier_reward(region.individual_tiers, qty, rev, distinct)

    def calc_store_bonus(self, project_id: str, branch: str) -> float:
        proj = self._get_project(project_id)
        region = proj.regions.get(self.region_of(branch))
        if not region:
            return 0.0
        qty, rev, distinct = self.get_branch_value(branch, project_id)
        return best_tier_reward(region.store_tiers, qty, rev, distinct)

    def all_sellers(self):
        """Trả về set các (seller, branch) đã từng phát sinh doanh số dự án mini."""
        seen = {}
        for (seller, _proj_id), d in self._seller_data.items():
            seen[seller] = d["branch"]
        return list(seen.items())

    def _get_project(self, project_id: str) -> MiniProject:
        for p in self.projects:
            if p.id == project_id:
                return p
        raise KeyError(f"Không tìm thấy dự án mini với id='{project_id}' trong config")


# --------------------------------------------------------------------------- #
# 4. Hàm tiện ích: quét file hóa đơn/trả hàng KiotViet (dùng chung mọi tháng)
# --------------------------------------------------------------------------- #

def find_col(header_map: Dict[str, int], keywords: List[str], default_idx: int) -> int:
    for k, idx in header_map.items():
        if any(kw in k for kw in keywords):
            return idx
    return default_idx


def scan_invoice_file_into_engine(engine: MiniProjectsEngine, inv_file: str,
                                   normalize_branch_fn):
    """
    Quét 1 file hóa đơn chi tiết KiotViet và đổ dữ liệu vào engine.
    Tự dò cột theo tên header (không hardcode vị trí cột), nên vẫn chạy được
    dù KiotViet đổi thứ tự cột giữa các tháng.
    """
    import openpyxl

    wb = openpyxl.load_workbook(inv_file, read_only=True, data_only=True)
    ws = wb.active
    it = ws.iter_rows(values_only=True)
    hdr = next(it)
    h_map = {str(h).strip().lower(): i for i, h in enumerate(hdr) if h}

    c_cn = find_col(h_map, ['chi nhánh', 'branch'], 0)
    c_seller = find_col(h_map, ['người bán', 'seller', 'nhân viên'], 7)
    c_sku = find_col(h_map, ['mã hàng', 'mã sản phẩm', 'sku'], 14)
    c_qty = find_col(h_map, ['số lượng', 'so_luong', 'quantity'], 21)
    c_tt = find_col(h_map, ['thành tiền', 'thanh_tien'], 26)

    n_rows = 0
    for r in it:
        if len(r) <= max(c_cn, c_seller, c_sku, c_qty, c_tt):
            continue
        branch = normalize_branch_fn(r[c_cn])
        seller = str(r[c_seller]).strip() if r[c_seller] else ''
        sku = str(r[c_sku]).strip() if r[c_sku] else ''
        qty = float(r[c_qty] or 0)
        tt = float(r[c_tt] or 0)
        if not sku or not seller:
            continue
        engine.add_sale(seller, branch, sku, qty, tt)
        n_rows += 1
    wb.close()
    print(f"--> [MiniProjectsEngine] Đã quét {n_rows} dòng hóa đơn liên quan dự án mini.")
    return engine
