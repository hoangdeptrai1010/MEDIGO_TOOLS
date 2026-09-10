# -*- coding: utf-8 -*-
"""
Đối chiếu MiniProjectsEngine (mới) với công thức Excel hardcode cũ (bpr_v2.py)
để đảm bảo refactor không làm SAI số tiền thưởng thực tế.

Công thức cũ (HN), trích từ bpr_v2.py:
  KAT (dòng 708):   IF(qty>=17, qty*30000, IF(qty>=15, qty*25000, 0))
  PartySmart(709):  IF(qty>=80, qty*6000, 0)
  Ladycare (710):   IF(rev>7000000, rev*0.06, IF(rev>5000000, rev*0.04, 0))
  AVC (711):        IF(rev_gate>9000000, qty*6000, IF(rev_gate>7000000, qty*4000, 0))

  Store bonus (dòng 767-779):
  KAT store:        IF(store_qty>=17, 300000, 0)          (bỏ qua điều kiện >=3 mã cho phép so sánh)
  PartySmart store: IF(store_qty>=180, 300000, 0)
  Ladycare store:   IF(store_rev>30000000, 700000, IF(store_rev>22000000, 500000, 0))
"""

from mini_projects_core import Tier, best_tier_reward

MINI_PROJECTS_CFG = [
    {
        "id": "kat", "name": "Mini KAT (Organika)",
        "regions": {
            "HN": {
                "skus": ["SP2725289", "SP2725285", "SP2725291", "SP2725287", "SP2725353", "SP2725351"],
                "individual_tiers": [
                    {"gate_on": "qty", "min": 15, "reward": {"type": "per_unit", "basis": "qty", "rate": 25000}},
                    {"gate_on": "qty", "min": 17, "reward": {"type": "per_unit", "basis": "qty", "rate": 30000}},
                ],
                "store_tiers": [
                    {"gate_on": "qty", "min": 17, "reward": {"type": "flat", "basis": "qty", "rate": 300000}},
                ],
            }
        },
    },
    {
        "id": "party_smart", "name": "Party Smart (Giải rượu Himalaya)",
        "regions": {
            "HN": {
                "skus": ["SP017162"],
                "individual_tiers": [
                    {"gate_on": "qty", "min": 80, "reward": {"type": "per_unit", "basis": "qty", "rate": 6000}},
                ],
                "store_tiers": [
                    {"gate_on": "qty", "min": 180, "reward": {"type": "flat", "basis": "qty", "rate": 300000}},
                ],
            }
        },
    },
    {
        "id": "ladycare", "name": "Ladycare (Bọt vệ sinh)",
        "regions": {
            "HN": {
                "skus": ["SP2723124", "SP2723125", "SP2723127"],
                "individual_tiers": [
                    {"gate_on": "revenue", "min": 5000000, "reward": {"type": "percent", "basis": "revenue", "rate": 0.04}},
                    {"gate_on": "revenue", "min": 7000000, "reward": {"type": "percent", "basis": "revenue", "rate": 0.06}},
                ],
                "store_tiers": [
                    {"gate_on": "revenue", "min": 22000000, "reward": {"type": "flat", "basis": "revenue", "rate": 500000}},
                    {"gate_on": "revenue", "min": 30000000, "reward": {"type": "flat", "basis": "revenue", "rate": 700000}},
                ],
            }
        },
    },
    {
        "id": "avc", "name": "AVC (Chất xơ, trùng thảo)",
        "regions": {
            "HN": {
                "skus": ["SP2723883", "SP2722826", "SP2722817", "SP2725666"],
                "individual_tiers": [
                    {"gate_on": "revenue", "min": 7000000, "reward": {"type": "per_unit", "basis": "qty", "rate": 4000}},
                    {"gate_on": "revenue", "min": 9000000, "reward": {"type": "per_unit", "basis": "qty", "rate": 6000}},
                ],
                "store_tiers": [],
            }
        },
    },
]


def old_formula_kat(qty):
    if qty >= 17: return qty * 30000
    if qty >= 15: return qty * 25000
    return 0

def old_formula_party_smart(qty):
    if qty >= 80: return qty * 6000
    return 0

def old_formula_ladycare(rev):
    if rev > 7000000: return rev * 0.06
    if rev > 5000000: return rev * 0.04
    return 0

def old_formula_avc(rev_gate, qty):
    if rev_gate > 9000000: return qty * 6000
    if rev_gate > 7000000: return qty * 4000
    return 0

def old_formula_kat_store(qty):
    return 300000 if qty >= 17 else 0

def old_formula_party_smart_store(qty):
    return 300000 if qty >= 180 else 0

def old_formula_ladycare_store(rev):
    if rev > 30000000: return 700000
    if rev > 22000000: return 500000
    return 0


def get_tiers(project_id, kind):
    for p in MINI_PROJECTS_CFG:
        if p["id"] == project_id:
            region_def = p["regions"]["HN"]
            key = "individual_tiers" if kind == "individual" else "store_tiers"
            return [Tier.from_dict(t) for t in region_def[key]]
    raise KeyError(project_id)


def run_tests():
    fails = 0

    # --- KAT individual ---
    kat_tiers = get_tiers("kat", "individual")
    for qty in [0, 10, 15, 16, 17, 20, 50]:
        new = best_tier_reward(kat_tiers, qty, 0)
        old = old_formula_kat(qty)
        ok = abs(new - old) < 0.01
        fails += (not ok)
        print(f"KAT qty={qty:>3}  new={new:>10.0f}  old={old:>10.0f}  {'OK' if ok else '<<< SAI'}")

    # --- Party Smart individual ---
    ps_tiers = get_tiers("party_smart", "individual")
    for qty in [0, 79, 80, 81, 200]:
        new = best_tier_reward(ps_tiers, qty, 0)
        old = old_formula_party_smart(qty)
        ok = abs(new - old) < 0.01
        fails += (not ok)
        print(f"PartySmart qty={qty:>3}  new={new:>10.0f}  old={old:>10.0f}  {'OK' if ok else '<<< SAI'}")

    # --- Ladycare individual (revenue) ---
    lc_tiers = get_tiers("ladycare", "individual")
    for rev in [0, 4_000_000, 5_000_001, 6_000_000, 7_000_001, 10_000_000]:
        new = best_tier_reward(lc_tiers, 0, rev)
        old = old_formula_ladycare(rev)
        ok = abs(new - old) < 0.01
        fails += (not ok)
        print(f"Ladycare rev={rev:>10}  new={new:>12.0f}  old={old:>12.0f}  {'OK' if ok else '<<< SAI'}")

    # --- AVC (gate = revenue, reward basis = qty) ---
    avc_tiers = get_tiers("avc", "individual")
    for rev, qty in [(6_000_000, 10), (7_000_001, 10), (9_000_001, 10), (0, 100)]:
        new = best_tier_reward(avc_tiers, qty, rev)
        old = old_formula_avc(rev, qty)
        ok = abs(new - old) < 0.01
        fails += (not ok)
        print(f"AVC rev={rev:>10} qty={qty:>3}  new={new:>10.0f}  old={old:>10.0f}  {'OK' if ok else '<<< SAI'}")

    # --- Store-level tiers ---
    kat_store = get_tiers("kat", "store")
    for qty in [16, 17, 30]:
        new = best_tier_reward(kat_store, qty, 0)
        old = old_formula_kat_store(qty)
        ok = abs(new - old) < 0.01
        fails += (not ok)
        print(f"KAT-store qty={qty:>3}  new={new:>10.0f}  old={old:>10.0f}  {'OK' if ok else '<<< SAI'}")

    ps_store = get_tiers("party_smart", "store")
    for qty in [179, 180, 300]:
        new = best_tier_reward(ps_store, qty, 0)
        old = old_formula_party_smart_store(qty)
        ok = abs(new - old) < 0.01
        fails += (not ok)
        print(f"PartySmart-store qty={qty:>3}  new={new:>10.0f}  old={old:>10.0f}  {'OK' if ok else '<<< SAI'}")

    lc_store = get_tiers("ladycare", "store")
    for rev in [20_000_000, 22_000_001, 25_000_000, 30_000_001]:
        new = best_tier_reward(lc_store, 0, rev)
        old = old_formula_ladycare_store(rev)
        ok = abs(new - old) < 0.01
        fails += (not ok)
        print(f"Ladycare-store rev={rev:>10}  new={new:>10.0f}  old={old:>10.0f}  {'OK' if ok else '<<< SAI'}")

    print()
    if fails == 0:
        print("=== TẤT CẢ TEST KHỚP 100% VỚI CÔNG THỨC CŨ ===")
    else:
        print(f"=== CÓ {fails} TRƯỜNG HỢP SAI LỆCH — CẦN XEM LẠI ===")
    return fails


if __name__ == "__main__":
    run_tests()
