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
from settlement.services.purchase_voucher import build_purchase_voucher
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
    # 按销售大类展开（与明细页默认一致）：贷方「库存商品—纱线」「发出商品—面料」都看得见
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
                "（结算后按销售数量 × 月末加权平均价回填）。借方 主营业务成本／发出商品、"
                "贷方 库存商品／发出商品，借贷两侧都按销售大类展开。",
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
    # 与销售凭证卡片同格式：按物料大类展开、贷方汇总一行，5 列（摘要｜科目｜数量｜借方｜贷方）
    target = _period_by_month(period, month)
    voucher = build_purchase_voucher(target, group="category")
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
        "note": "本月到货本月到票直接入账；到货未到票暂估；前期暂估本月到票回冲原暂估并按票入账；"
                "价差进材料成本差异。期初暂估台账余额只提示、不入账。",
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


def build_voucher_center(period, sales_month=None, purchase_month=None):
    """本期凭证中心数据：已实现 2 类 + 规划中 3 类。

    sales_month / purchase_month：两张卡片各自「凭证期间」的筛选（选哪个期间就出哪个期间的凭证）。
    """
    # 按建议的生成顺序排列：暂估回冲 → 投料 → 制造费用 → 完工入库 → 销售结转
    cards = [_purchase_card(period, month=purchase_month)]
    cards.append(
        _planned(
            "material_issue",
            "生产领料（投料）凭证",
            "MATERIAL ISSUE VOUCHER",
            "借 生产成本—直接材料（按车间）/ 贷 原材料。取「库存出库流水」口径"
            "（投料 − 投料回撤），不能用工单直接材料（含自动完工补料与车间内部流转）。"
            "无工单领料（浆粕原料等）的去向需先确认。",
        )
    )
    cards.append(
        _planned(
            "manufacturing_overhead",
            "制造费用 / 人工结转凭证",
            "OVERHEAD TRANSFER VOUCHER",
            "借 生产成本—制造费用、直接人工 / 贷 制造费用（水电、蒸汽、天然气、折旧、"
            "机物料、染料助剂）、应付职工薪酬。取车间费用池。",
        )
    )
    cards.append(
        _planned(
            "production_receipt",
            "完工入库结转凭证",
            "FINISHED GOODS TRANSFER VOUCHER",
            "借 库存商品 / 自制半成品 / 贷 生产成本。金额基准（库存入库流水 或 "
            "工单完工成本）与浆粕单独归集需先确认；本地 7 月两条口径相差 362.57 万。",
        )
    )
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
