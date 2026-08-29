# -*- coding: utf-8 -*-
"""浆粕车间成本归集（用户 2026-08-28 定稿口径）。

浆粕是特殊车间：无生产工单（PULP202607 占位）、无投料记录、无报工。
成本归集 = 原料领料(用户上传"浆粕原料流水") + 车间费用，按产量分摊到 4 种成品，
并回填"生产入库"流水单价（完工成本）。

口径：
- 原料识别：用户上传的「浆粕原料流水」（材料成本计算来源，未上传不能结转）
  - 名称含"芳纶" → 芳纶类原料
  - 名称含"腈"   → 腈纶类原料
- 成品识别：物料名含"浆粕"或"微纤"，或 JP- 系列（与导出中心一致）
- 芳纶类成品：21.03 芳纶浆粕 DFP-103、22.01 芳纶微纤 AMF-100（名含芳纶）
- 腈纶类成品：23.06 腈纶浆粕 JLP-505、24.01 腈纶微纤 JMF-025（名含腈）
- 材料分摊：同类原料金额 → 同类成品按 生产入库产量比例
- 费用/工资：车间费用(水电/蒸汽/天然气/折旧/机物料/染料助剂/工资) → 全部成品按 生产入库产量比例
- 回填：各成品"生产入库"流水 单价=完工成本/产量、金额=完工成本
- 产量=0 的成品：材料/费用都分不到（金额 0，单价 0）
"""
from __future__ import annotations

from collections import defaultdict
from decimal import Decimal

from settlement.models import AccountingPeriod, InventoryTransaction

ZERO = Decimal("0")
CENT = Decimal("0.01")

FEE_FIELDS = (
    "water_power",
    "steam",
    "natural_gas",
    "depreciation",
    "consumables",
    "dye_auxiliary",
    "wages",
)




def _is_pulp_product(code, name) -> bool:
    return ("浆粕" in str(name or "")) or ("微纤" in str(name or "")) or str(code or "").upper().startswith("JP-")


def _is_fanglun(name) -> bool:
    return "芳纶" in str(name or "")


def _is_jinglun(name) -> bool:
    return "腈" in str(name or "")


def _material_price(period, code) -> Decimal | None:
    """浆粕原料出库单价：历史单价出库价优先、入库价兜底（与投料定价口径一致）。"""
    hp = period.historical_prices.filter(material_code=code).order_by("-id").first()
    if not hp:
        return None
    if hp.issue_price:
        return Decimal(hp.issue_price)
    if hp.receipt_price:
        return Decimal(hp.receipt_price)
    return None



def collect_pulp_costs(period: AccountingPeriod) -> dict:
    """浆粕成本归集，返回 {"products": {成品编码: {"name","output","material","fee","wages","total","unit"}},
                                  "material_fanglun": 金额, "material_jinglun": 金额, "fee_total": 金额,
                                  "output_total": 产量合计}"""
    # ---- 1) 原料领料（用户上传的浆粕原料流水，材料成本计算来源）----
    #      只认业务类型为"领料出库"的行；金额缺省时按历史单价（出库优先、入库兜底）计算
    mat_fanglun = ZERO   # 芳纶类原料金额
    mat_jinglun = ZERO   # 腈纶类原料金额
    issue_detail = []
    uploaded = list(period.pulp_materials.all().order_by("id"))
    if not uploaded:
        raise ValueError(
            "未上传浆粕原料流水：请先在「资料导入」页上传当月浆粕原料流水，再结转浆粕车间。"
        )
    for m in uploaded:
        if "领料出库" not in (m.business_type or ""):
            continue  # 非领料出库（如生产入库/销售出库）不计入材料成本
        name = m.material_name or ""
        if not (_is_fanglun(name) or _is_jinglun(name)):
            raise ValueError(
                f"浆粕原料流水第{m.id}行（{m.material_code} {name}）无法归类："
                "物料名称须含「芳纶」或「腈」，请修正文件后重新上传。"
            )
        amt = Decimal(m.amount or ZERO)
        if amt <= 0:
            price = _material_price(period, m.material_code)
            if price:
                amt = (Decimal(m.quantity or ZERO) * price).quantize(CENT)
            else:
                raise ValueError(
                    f"浆粕原料流水{m.material_code} {name}没有金额，且查不到历史单价"
                    "（出库/入库单价），请上传历史单价或在该行填写金额。"
                )
        if _is_fanglun(name):
            mat_fanglun += amt
        else:
            mat_jinglun += amt
        issue_detail.append((m.document_no, m.material_code, name, m.quantity, amt))

    # ---- 2) 成品与产量（生产入库数量）----
    products = {}  # code -> {"name", "category", "output"}
    prod_qty = defaultdict(Decimal)
    prod_amt_old = defaultdict(Decimal)
    for t in period.inventory_transactions.filter(business_type="生产入库"):
        if _is_pulp_product(t.material_code, t.material_name):
            products.setdefault(t.material_code, {
                "name": t.material_name,
                "category": "芳纶" if _is_fanglun(t.material_name) else ("腈纶" if _is_jinglun(t.material_name) else "其他"),
            })
            prod_qty[t.material_code] += Decimal(t.quantity or ZERO)
            prod_amt_old[t.material_code] += Decimal(t.amount or ZERO)

    # ---- 3) 车间费用 ----
    pool = period.workshop_cost_pools.filter(workshop="浆粕车间").first()
    fee_by_field = {f: (Decimal(getattr(pool, f, ZERO) or ZERO) if pool else ZERO) for f in FEE_FIELDS}
    mfg_total = sum((v for f, v in fee_by_field.items() if f != "wages"), ZERO)
    wages_total = fee_by_field.get("wages", ZERO)

    output_total = sum(prod_qty.values(), ZERO)
    # 同类产量合计（材料按同类内产量比例分摊）
    cat_output = defaultdict(Decimal)
    for code, info in products.items():
        cat_output[info["category"]] += prod_qty[code]

    # ---- 4) 分摊 ----
    result = {}
    for code, info in products.items():
        qty = prod_qty[code]
        cat = info["category"]
        if cat == "芳纶":
            cat_tot = cat_output.get("芳纶", ZERO)
            mat_share = mat_fanglun * (qty / cat_tot) if cat_tot else ZERO
        elif cat == "腈纶":
            cat_tot = cat_output.get("腈纶", ZERO)
            mat_share = mat_jinglun * (qty / cat_tot) if cat_tot else ZERO
        else:
            mat_share = ZERO
        fee_share = mfg_total * (qty / output_total) if output_total else ZERO
        wage_share = wages_total * (qty / output_total) if output_total else ZERO
        total = mat_share + fee_share + wage_share
        unit = (total / qty).quantize(Decimal("0.0001")) if qty else ZERO
        result[code] = {
            "name": info["name"],
            "category": cat,
            "output": qty,
            "material": mat_share,
            "fee": fee_share,
            "wages": wage_share,
            "total": total,
            "unit": unit,
        }

    return {
        "products": result,
        "material_fanglun": mat_fanglun,
        "material_jinglun": mat_jinglun,
        "mfg_total": mfg_total,
        "wages_total": wages_total,
        "output_total": output_total,
        "issue_detail": issue_detail,
    }


def price_pulp_receipts(period: AccountingPeriod) -> Decimal:
    """回填浆粕成品"生产入库"流水单价/金额 = 完工成本，返回最大变动金额。"""
    costs = collect_pulp_costs(period)
    changes = []
    for code, c in costs["products"].items():
        if c["output"] <= ZERO:
            continue
        unit = c["unit"]
        amt = money(c["total"])
        for t in period.inventory_transactions.filter(
            business_type="生产入库", material_code=code
        ):
            old = Decimal(t.amount or ZERO)
            t.unit_price = unit
            t.amount = amt
            t.price_source = "浆粕完工成本"
            t.save(update_fields=["unit_price", "amount", "price_source"])
            changes.append(abs(Decimal(t.amount) - old))
    return max(changes, default=ZERO)


def money(value) -> Decimal:
    from decimal import ROUND_HALF_UP
    return Decimal(value or ZERO).quantize(CENT, rounding=ROUND_HALF_UP)
