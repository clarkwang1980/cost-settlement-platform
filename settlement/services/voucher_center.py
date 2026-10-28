# -*- coding: utf-8 -*-
"""凭证中心：把本期各类成本凭证列在一个页面上。

每类凭证给出：关键金额、借贷合计、前置检查、分录明细（只读）、以及打开明细页的入口。
已实现的：销售结转成本凭证、采购暂估/回冲凭证；
规划中的：生产领料（投料）、制造费用结转、完工入库结转（见《成本凭证模块_讨论稿.md》）。
"""
from __future__ import annotations

from datetime import date
from decimal import Decimal

from settlement.models import AccountingPeriod
from settlement.services.material_issue_voucher import build_material_issue_voucher
from settlement.services.receipt_voucher import build_finished_receipt_voucher
from settlement.services.purchase_contract_voucher import build_invoice_matched_voucher
from settlement.services.sales_voucher import (
    CARRY_COST_TYPES,
    build_sales_voucher_entries,
)

ZERO = Decimal("0")


def _metric(label, value, tone=""):
    return {"label": label, "value": value, "tone": tone}


def _cell(text, num=False):
    return {"text": text, "num": num}


def _sales_card(period, month=None):
    # 与明细页默认一致（group=category）：贷方按销售大类展开看得到「库存商品—纱线」，
    # 借方「主营业务成本」只到销售渠道（国内/国外）
    target = None
    if month:
        try:
            year, month_no = str(month).split("-")
            target = date(int(year), int(month_no), 1)
        except (TypeError, ValueError):
            target = None
    voucher = build_sales_voucher_entries(period, month=target, group="category")
    totals = voucher["totals"]
    entries = []
    for block in voucher["blocks"]:
        for row in block["entries"]:
            entries.append(
                {
                    "level": row["level"],
                    "cells": [
                        _cell(row["summary"]),
                        _cell(row["account"]),
                        _cell(
                            "{:,.3f}".format(row["quantity"]) if row["quantity"] else "",
                            True,
                        ),
                        _cell("{:,.2f}".format(row["debit"]) if row["debit"] else "", True),
                        _cell("{:,.2f}".format(row["credit"]) if row["credit"] else "", True),
                    ],
                }
            )
    checks = []
    if totals["missing_cost_rows"]:
        checks.append(
            {
                "level": "warning",
                "text": "有 %d 行销售台账没有销售成本（该产品当月没有库存计价），"
                        "这些行的金额不含在内。" % totals["missing_cost_rows"],
            }
        )
    if not totals["source_rows"]:
        checks.append(
            {
                "level": "warning",
                "text": "凭证期间 %s 没有参与结转的销售台账行。" % voucher["month"],
            }
        )
    return {
        "key": "sales",
        "title": "销售结转成本凭证",
        "eyebrow": "SALES COST CARRY-FORWARD VOUCHER",
        "status": "ready",
        "status_label": "已实现",
        "note": "数据来源：销售台账。结转类型与销售大类都取台账原值；金额取台账「销售成本」列"
                "（结算后按销售数量 × 月末加权平均价回填）。"
                "借方 主营业务成本（只分国内/国外销售）／发出商品、"
                "贷方 库存商品／发出商品（按销售大类展开）。",
        "url_name": "sales_cost_voucher",
        "export_kind": "sales_voucher",
        # 汇总字段按要求去掉；凭证期间保留，并做成可筛选字段（卡片上的下拉）
        "metrics": [],
        "month": voucher["month"],
        "month_filter": {
            "param": "sales_month",
            "value": voucher["month"],
            "options": [
                {"value": item["value"], "label": item["label"]}
                for item in voucher["months"]
            ],
        },
        "entry_headers": list(voucher["headers"]),
        "entries": entries,
        "balance": None,
        "checks": checks,
    }



def _period_by_month(period, month):
    """按 YYYY-MM 找核算期间；找不到就用当前期间。"""
    if month:
        try:
            year, month_no = str(month).split("-")
            found = AccountingPeriod.objects.filter(
                month=date(int(year), int(month_no), 1)
            ).first()
            if found:
                return found
        except (TypeError, ValueError):
            pass
    return period


def _purchase_card(period, month=None):
    # 与销售凭证卡片同格式：原材料与应付账款都汇总一行，5 列（摘要｜科目｜数量｜借方｜贷方）
    target = _period_by_month(period, month)
    voucher = build_invoice_matched_voucher(target)
    entries = []
    for block in voucher["blocks"]:
        for row in block["rows"]:
            entries.append(
                {
                    "level": row["level"],
                    "cells": [
                        _cell(row["summary"]),
                        _cell(row["account"]),
                        _cell(
                            "{:,.3f}".format(row["quantity"]) if row["quantity"] else "",
                            True,
                        ),
                        _cell("{:,.2f}".format(row["debit"]) if row["debit"] else "", True),
                        _cell("{:,.2f}".format(row["credit"]) if row["credit"] else "", True),
                    ],
                }
            )
    return {
        "key": "purchase",
        "title": "采购暂估 / 回冲凭证",
        "eyebrow": "PURCHASE ACCRUAL VOUCHER",
        "status": "ready",
        "status_label": "已实现",
        "note": "到货＝期初暂估台账 ＋ 合同「已签收数量」；到票＝采购发票对照合同/台账核销。"
                "本期发票冲本期到货→到货到票；冲以前期间到货→回冲（借 暂估应付／贷 应付账款）；"
                "本期到货未到票→暂估（借 原材料／贷 暂估应付）；价差进材料成本差异。",
        "url_name": "purchase_accrual_voucher",
        "export_kind": "purchase_voucher",
        # 汇总字段与借贷 pill 与销售卡片一致，不再展示；凭证期间做成筛选字段
        "metrics": [],
        "month": voucher["month"],
        "month_filter": {
            "param": "purchase_month",
            "value": voucher["month"],
            "options": [
                {"value": item.month.strftime("%Y-%m"), "label": str(item)}
                for item in AccountingPeriod.objects.order_by("month")
            ],
        },
        "entry_headers": list(voucher["headers"]),
        "entries": entries,
        "balance": None,
        "checks": voucher["checks"],
    }


def _material_issue_card(period, month=None):
    """生产领料（投料）凭证卡片：与采购/销售卡片同格式（5 列、无汇总指标）。"""
    target = _period_by_month(period, month)
    voucher = build_material_issue_voucher(target)
    entries = [
        {
            "level": row["level"],
            "cells": [
                _cell(row["summary"]),
                _cell(row["account"]),
                _cell(
                    "{:,.3f}".format(row["quantity"]) if row["quantity"] else "",
                    True,
                ),
                _cell("{:,.2f}".format(row["debit"]) if row["debit"] else "", True),
                _cell("{:,.2f}".format(row["credit"]) if row["credit"] else "", True),
            ],
        }
        for row in voucher["rows"]
    ]
    note = (
        "一张凭证、不分块：借方科目合并——各车间合计成一行「生产成本—直接材料」；"
        "贷方按车间顺序（染色→纺纱→包覆→针织→机织→浆粕→无工单领料）逐行列出各车间领用的物料大类。"
        "取数：除浆粕车间外按「工单投料台账」（实际投料数量×月末加权平均价，工单→生产部门归车间）；"
        "浆粕车间没有工单，按库存流水「领料出库」且备注含「浆粕 + 当期月份」认领；"
        "其余无工单领料单独列示。"
    )
    return {
        "key": "material_issue",
        "title": "生产领料（投料）凭证",
        "eyebrow": "MATERIAL ISSUE VOUCHER",
        "status": "ready",
        "status_label": "已实现",
        "note": note,
        "url_name": "material_issue_voucher",
        "export_kind": "material_issue_voucher",
        "metrics": [],
        "month": voucher["month"],
        "month_filter": {
            "param": "material_issue_month",
            "value": voucher["month"],
            "options": [
                {"value": item.month.strftime("%Y-%m"), "label": str(item)}
                for item in AccountingPeriod.objects.order_by("month")
            ],
        },
        "entry_headers": list(voucher["headers"]),
        "entries": entries,
        "balance": None,
        "checks": voucher["checks"],
    }


def _production_receipt_card(period, month=None):
    """完工入库结转凭证卡片（用户 2026-09-25；贷方改成本要素 2026-09-27）：

    摘要按车间、科目按物料分类 —— 借 库存商品—物料分类；
    贷方 生产成本—直接材料 / 直接人工 / 制造费用（全车间汇总，含浆粕车间）。
    金额基准＝工单完工成本；卡片上不再放借贷金额 pill 与附件 1/2，明细统一到「打开明细」页。
    """
    target = _period_by_month(period, month)
    voucher = build_finished_receipt_voucher(target)
    entries = [
        {
            "level": row["level"],
            "cells": [
                _cell(row["summary"]),
                _cell(row["account"]),
                _cell(
                    "{:,.3f}".format(row["quantity"]) if row["quantity"] else "",
                    True,
                ),
                _cell("{:,.2f}".format(row["debit"]) if row["debit"] else "", True),
                _cell("{:,.2f}".format(row["credit"]) if row["credit"] else "", True),
            ],
        }
        for row in voucher["rows"]
    ]
    summary = voucher["summary"]
    note = (
        "摘要按车间、科目按物料分类：借方 库存商品—物料分类（每个车间的每个物料分类一行，"
        "摘要「{月}月{车间}完工入库{物料分类}」）；贷方 生产成本—直接材料 / 直接人工 / 制造费用"
        "（全车间汇总，各一行，摘要「{月}月各车间完工入库{要素}」）；浆粕车间没有工单，"
        "按库存流水「生产入库」归集后**也拆成这三个要素**。"
        "金额基准＝**工单完工成本**（报工数量 × 完工单位成本，与「工单成本」页一致）："
        "本期 %d 张工单 %s 元（直接材料 %s / 直接人工 %s / 制造费用 %s）；"
        "浆粕车间归集 %d 个成品 %s 元（直接材料 %s / 直接人工 %s / 制造费用 %s）。"
        % (summary["order_count"], "{:,.2f}".format(summary["order_amount"]),
           "{:,.2f}".format(summary["order_element_totals"]["直接材料"]),
           "{:,.2f}".format(summary["order_element_totals"]["直接人工"]),
           "{:,.2f}".format(summary["order_element_totals"]["制造费用"]),
           summary["pulp_count"], "{:,.2f}".format(summary["pulp_amount"]),
           "{:,.2f}".format(summary["pulp_element_totals"]["直接材料"]),
           "{:,.2f}".format(summary["pulp_element_totals"]["直接人工"]),
           "{:,.2f}".format(summary["pulp_element_totals"]["制造费用"]))
    )
    return {
        "key": "production_receipt",
        "title": "完工入库结转凭证",
        "eyebrow": "FINISHED GOODS TRANSFER VOUCHER",
        "status": "ready",
        "status_label": "已实现",
        "note": note,
        "url_name": "finished_receipt_voucher",
        "export_kind": "finished_receipt_voucher",
        "metrics": [],
        "month": voucher["month"],
        "month_filter": {
            "param": "receipt_month",
            "value": voucher["month"],
            "options": [
                {"value": item.month.strftime("%Y-%m"), "label": str(item)}
                for item in AccountingPeriod.objects.order_by("month")
            ],
        },
        "entry_headers": list(voucher["headers"]),
        "entries": entries,
        # 用户 2026-09-27：卡片不再显示借贷金额 pill；明细（原附件1/2）搬到「打开明细」页
        "balance": None,
        "checks": voucher["checks"],
    }


def _planned(key, title, eyebrow, note):
    return {
        "key": key,
        "title": title,
        "eyebrow": eyebrow,
        "status": "planned",
        "status_label": "规划中",
        "note": note,
        "url_name": "",
        "export_kind": "",
        "metrics": [],
        "entry_headers": [],
        "entries": [],
        "balance": None,
        "checks": [],
    }


def build_voucher_center(period, sales_month=None, purchase_month=None,
                         material_issue_month=None, receipt_month=None):
    """本期凭证中心数据：已实现 4 类 + 规划中 1 类。

    sales_month / purchase_month / material_issue_month / receipt_month：
    各卡片自己的「凭证期间」筛选（选哪个期间就出哪个期间的凭证）。
    """
    # 按建议的生成顺序排列：暂估回冲 → 投料 → 制造费用 → 完工入库 → 销售结转
    cards = [_purchase_card(period, month=purchase_month)]
    cards.append(_material_issue_card(period, month=material_issue_month))
    cards.append(
        _planned(
            "manufacturing_overhead",
            "制造费用 / 人工结转凭证",
            "OVERHEAD TRANSFER VOUCHER",
            "借 生产成本—制造费用、直接人工 / 贷 制造费用（水电、蒸汽、天然气、折旧、"
            "机物料、染料助剂）、应付职工薪酬。取车间费用池。",
        )
    )
    cards.append(_production_receipt_card(period, month=receipt_month))
    cards.append(_sales_card(period, month=sales_month))

    # 顺序号：页面按这个顺序展示，末尾的「建议的凭证生成顺序」也用它
    for index, card in enumerate(cards, start=1):
        card["step"] = index

    ready = [card for card in cards if card["status"] == "ready"]
    blocking = [
        check
        for card in cards
        for check in card["checks"]
        if check["level"] == "blocking"
    ]
    warnings = [
        check
        for card in cards
        for check in card["checks"]
        if check["level"] == "warning"
    ]
    return {
        "month": period.month.strftime("%Y-%m"),
        "period_status": period.get_status_display(),
        "locked": bool(period.locked_at),
        "cards": cards,
        "steps": [
            {
                "step": card["step"],
                "title": card["title"],
                "key": card["key"],
                "status": card["status"],
                "status_label": card["status_label"],
            }
            for card in cards
        ],
        "ready_count": len(ready),
        "planned_count": len(cards) - len(ready),
        "entry_count": sum(len(card["entries"]) for card in cards),
        "blocking": blocking,
        "warnings": warnings,
        "all_balanced": all(
            card["balance"]["balanced"] for card in cards if card["balance"]
        ),
    }
