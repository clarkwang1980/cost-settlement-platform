# -*- coding: utf-8 -*-
"""总成本表数据汇总（页面 + 导出共用）。

口径（与手工 Excel 生成脚本一致）：
- 行 = 科目：直接材料 / 水电费 / 蒸汽费 / 燃气费 / 染料助剂费 / 机物料费 / 折旧费 / 人工费
- 列 = 四大块：期初 / 本期投入 / 产成品成本 / 在产品；每块 = 各车间 + 合计
- 期初     = 期初在产品表(OpeningWip)按工单车间汇总（8 科目）
- 本期投入 = 工单结算结果(WorkOrderCostResult)按车间汇总(8 科目) − 期初
- 产成品/在产品 = 结算结果 finished/wip 拆分：材料、人工直接取；
  制造费用（水电/蒸汽/燃气/染料助剂/机物料/折旧）按该车间制造费用结构比例拆回 6 科目
- 未结算车间（无结算结果）：本期投入/产成品/在产品 无数据（None -> 页面显示 --）
- 勾稽：期初 + 本期投入 == 产成品成本 + 在产品（对已结算车间逐科目校验）
"""
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP

from settlement.models import OpeningWip, WorkOrderCostResult

ZERO = Decimal("0.00")
CENT = Decimal("0.01")

SUBJECTS = ["直接材料", "水电费", "蒸汽费", "燃气费", "染料助剂费",
            "机物料费", "折旧费", "人工费"]
RESULT_COLS = ["direct_material", "water_power", "steam", "natural_gas",
               "dye_auxiliary", "consumables", "depreciation", "wages"]
OPENING_COLS = ["material", "water_power", "steam", "natural_gas",
                "dye_auxiliary", "consumables", "depreciation", "wages"]
MFG_FIELDS = ["water_power", "steam", "natural_gas", "dye_auxiliary",
              "consumables", "depreciation"]
MFG_SUBJECT = {"water_power": "水电费", "steam": "蒸汽费", "natural_gas": "燃气费",
               "dye_auxiliary": "染料助剂费", "consumables": "机物料费",
               "depreciation": "折旧费"}
BLOCKS = ["期初", "本期投入", "产成品成本", "在产品"]

DEFAULT_WORKSHOPS = ["RanZheng", "tefang", "BaoFu", "ZhenZhi", "JiZhi", "浆粕车间"]
WORKSHOP_LABELS = {"RanZheng": "染色车间", "tefang": "纺纱车间", "BaoFu": "包覆车间",
                   "ZhenZhi": "针织车间", "JiZhi": "机织车间", "浆粕车间": "浆粕车间"}


def _money(value) -> Decimal:
    return Decimal(value or ZERO).quantize(CENT, rounding=ROUND_HALF_UP)


def _workshop_order(period):
    """车间顺序：车间结算 sequence 优先，兜底固定顺序。"""
    from settlement.models import WorkshopSettlement
    seq = list(
        WorkshopSettlement.objects.filter(period=period)
        .order_by("sequence", "id")
        .values_list("workshop", flat=True)
    )
    order = [w for w in seq if w in WORKSHOP_LABELS]
    order += [w for w in DEFAULT_WORKSHOPS if w not in order]
    return order


def build_total_cost_data(period):
    """返回页面/导出共用的结构化数据。"""
    workshops = _workshop_order(period)
    ws_label = {code: WORKSHOP_LABELS.get(code, code) for code in workshops}

    # 浆粕车间：走专门成本归集（pulp_costing），无 WorkOrderCostResult。
    # 用归集结果填充：本期投入（材料=原料领料、制造费用=车间费用、人工=车间工资）与
    # 产成品成本（= 本期投入，浆粕无在产品、全完工）。
    PULP_WS = "浆粕车间"
    pulp_data = None
    if PULP_WS in workshops:
        from settlement.services.pulp_costing import collect_pulp_costs
        try:
            pulp_data = collect_pulp_costs(period)
        except Exception:
            pulp_data = None

    # 工单主档：order_no -> 车间
    wo_dept = dict(period.work_orders.values_list("order_no", "production_dept_code"))

    # 期初在产品按车间汇总（8 科目）
    opening = defaultdict(lambda: defaultdict(Decimal))
    for item in period.opening_wip.all():
        code = wo_dept.get(item.order_no) or ""
        for i, col in enumerate(OPENING_COLS):
            # 先精确累加，最后统一 quantize，避免逐行舍入累计误差
            opening[code][SUBJECTS[i]] += Decimal(getattr(item, col) or ZERO)

    # 结算结果按车间汇总（8 科目 + finished/wip）
    result = defaultdict(lambda: defaultdict(Decimal))
    results = list(
        WorkOrderCostResult.objects.filter(period=period)
        .select_related("work_order")
    )
    for r in results:
        code = r.work_order.production_dept_code or ""
        for i, col in enumerate(RESULT_COLS):
            result[code][SUBJECTS[i]] += Decimal(getattr(r, col) or ZERO)
        result[code]["finished_material"] += Decimal(r.finished_material or ZERO)
        result[code]["finished_wages"] += Decimal(r.finished_wages or ZERO)
        result[code]["finished_overhead"] += Decimal(r.finished_overhead or ZERO)
        result[code]["wip_material"] += Decimal(r.wip_material or ZERO)
        result[code]["wip_wages"] += Decimal(r.wip_wages or ZERO)
        result[code]["wip_overhead"] += Decimal(r.wip_overhead or ZERO)

    # 浆粕车间：把归集结果并入 result（材料/费用/工资科目 + finished 全量、wip=0）
    if pulp_data and pulp_data.get("products"):
        mfg_by_field = defaultdict(Decimal)
        # 用车间费用表结构拆分制造费用科目
        pool = period.workshop_cost_pools.filter(workshop=PULP_WS).first()
        if pool:
            for f in ("water_power", "steam", "natural_gas",
                      "dye_auxiliary", "consumables", "depreciation"):
                mfg_by_field[SUBJECTS[RESULT_COLS.index(f)]] = Decimal(getattr(pool, f) or ZERO)
        mat_total = Decimal(pulp_data["material_fanglun"] + pulp_data["material_jinglun"])
        wages_total = Decimal(pulp_data["wages_total"])
        result[PULP_WS]["直接材料"] += mat_total
        result[PULP_WS]["人工费"] += wages_total
        for subj, v in mfg_by_field.items():
            result[PULP_WS][subj] += v
        finished_total = sum(
            (Decimal(c["total"]) for c in pulp_data["products"].values()), ZERO
        )
        result[PULP_WS]["finished_material"] += mat_total
        result[PULP_WS]["finished_wages"] += wages_total
        result[PULP_WS]["finished_overhead"] += Decimal(pulp_data["mfg_total"])
        # wip 为 0（浆粕全完工）

    settled_codes = set(result.keys())
    # 汇总数据 matrix[block][code][subject]
    matrix = {block: {code: {s: None for s in SUBJECTS} for code in workshops}
              for block in BLOCKS}

    # 期初
    for code in workshops:
        for s in SUBJECTS:
            matrix["期初"][code][s] = opening[code].get(s, ZERO).quantize(
                CENT, rounding=ROUND_HALF_UP)

    # 本期投入 = 结果 − 期初（未结算车间为 None）
    for code in workshops:
        if code in settled_codes:
            for s in SUBJECTS:
                matrix["本期投入"][code][s] = (
                    result[code][s] - opening[code].get(s, ZERO)
                ).quantize(CENT, rounding=ROUND_HALF_UP)

    # 产成品 / 在产品
    for code in workshops:
        if code not in settled_codes:
            continue
        r = result[code]
        fin_mat = r["finished_material"]
        fin_wage = r["finished_wages"]
        fin_ovh = r["finished_overhead"]
        wip_mat = r["wip_material"]
        wip_wage = r["wip_wages"]
        wip_ovh = r["wip_overhead"]
        matrix["产成品成本"][code]["直接材料"] = fin_mat.quantize(
            CENT, rounding=ROUND_HALF_UP)
        matrix["产成品成本"][code]["人工费"] = fin_wage.quantize(
            CENT, rounding=ROUND_HALF_UP)
        matrix["在产品"][code]["直接材料"] = wip_mat.quantize(
            CENT, rounding=ROUND_HALF_UP)
        matrix["在产品"][code]["人工费"] = wip_wage.quantize(
            CENT, rounding=ROUND_HALF_UP)
        mfg_total = sum(r[MFG_SUBJECT[f]] for f in MFG_FIELDS)
        if mfg_total > 0:
            ratios = {f: r[MFG_SUBJECT[f]] / mfg_total for f in MFG_FIELDS}
        else:
            ratios = {f: ZERO for f in MFG_FIELDS}
        for block, total in (("产成品成本", fin_ovh), ("在产品", wip_ovh)):
            assigned = ZERO
            for i, f in enumerate(MFG_FIELDS):
                subj = MFG_SUBJECT[f]
                if i < len(MFG_FIELDS) - 1:
                    v = (total * ratios[f]).quantize(CENT, rounding=ROUND_HALF_UP)
                else:
                    v = (total - assigned).quantize(CENT, rounding=ROUND_HALF_UP)
                matrix[block][code][subj] = v
                assigned += v

    # 每块合计列 = 有值车间加总
    for block in BLOCKS:
        for s in SUBJECTS:
            vals = [matrix[block][code][s] for code in workshops
                    if matrix[block][code][s] is not None]
            matrix[block]["合计"] = matrix[block].get("合计", {})
            matrix[block]["合计"][s] = sum(vals).quantize(CENT) if vals else None

    # 行数据（模板友好）：每科目一行，cells 扁平 = 车间列 + 合计列
    rows = []
    for s in SUBJECTS:
        cells = []
        for block in BLOCKS:
            for code in workshops:
                cells.append({"value": matrix[block][code][s], "total": False})
            cells.append({"value": matrix[block]["合计"][s], "total": True})
        rows.append({"subject": s, "cells": cells})

    # 合计行（每列纵向科目加总）
    total_cells = []
    for block in BLOCKS:
        for code in workshops:
            vals = [matrix[block][code][s] for s in SUBJECTS
                    if matrix[block][code][s] is not None]
            total_cells.append({"value": sum(vals).quantize(CENT) if vals else None,
                                "total": False})
        vals = [matrix[block]["合计"][s] for s in SUBJECTS
                if matrix[block]["合计"][s] is not None]
        total_cells.append({"value": sum(vals).quantize(CENT) if vals else None,
                            "total": True})

    # 勾稽校验（已结算车间逐科目）
    balanced = True
    imbalance = []
    for code in settled_codes:
        for s in SUBJECTS:
            lhs = matrix["期初"][code][s] + matrix["本期投入"][code][s]
            rhs = matrix["产成品成本"][code][s] + matrix["在产品"][code][s]
            if lhs != rhs:
                balanced = False
                imbalance.append((code, s, lhs, rhs))

    pending = [code for code in workshops if code not in settled_codes]
    return {
        "period": period,
        "workshops": [{"code": c, "label": ws_label[c]} for c in workshops],
        "subjects": SUBJECTS,
        "blocks": BLOCKS,
        "rows": rows,
        "total_cells": total_cells,
        "settled_codes": settled_codes,
        "pending": pending,
        "balanced": balanced,
        "imbalance": imbalance,
    }
