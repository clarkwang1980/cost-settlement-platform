# -*- coding: utf-8 -*-
"""凭证中心：把本期各类成本凭证列在一个页面上。

每类凭证给出：关键金额、借贷合计、前置检查、分录明细（只读）、以及打开明细页的入口。
已实现的：销售结转成本凭证、采购暂估/回冲凭证；
规划中的：生产领料（投料）、制造费用结转、完工入库结转（见《成本凭证模块_讨论稿.md》）。
"""
from __future__ import annotations

from decimal import Decimal

from settlement.services.purchase_voucher import build_purchase_voucher
from settlement.services.sales_voucher import (
    CARRY_COST_TYPES,
    build_sales_voucher_summary,
)

ZERO = Decimal("0")


def _metric(label, value, tone=""):
    return {"label": label, "value": value, "tone": tone}


def _cell(text, num=False):
    return {"text": text, "num": num}


def _sales_card(period):
    voucher = build_sales_voucher_summary(period)
    totals = voucher["totals"]
    status_labels = {
        item["value"]: item["label"] for item in voucher["status_options"]
    }
    entries = []
    for row in voucher["rows"]:
        # 与凭证页保持完全一致的 8 列（含三个筛选列，这里以当前取值文本展示）
        if row["block_index"] >= 0:
            out_text = row["block_out"] or "（空白）"
            status_text = status_labels.get(
                row["block_status"], row["block_status"] or "（空白）"
            )
            invoice_text = row["block_invoice"] or "（空白）"
        else:
            out_text = status_text = invoice_text = ""
        entries.append(
            {
                "level": row["level"],
                "cells": [
                    _cell(row["account"]),
                    _cell(row["carry_type"]),
                    _cell(row["category"]),
                    _cell(out_text),
                    _cell(status_text),
                    _cell(invoice_text),
                    _cell("{:,.3f}".format(row["quantity"]), True),
                    _cell("{:,.2f}".format(row["amount"]), True),
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
                "text": "记账月份 %s 没有参与汇总的销售台账行。" % voucher["month"],
            }
        )
    return {
        "key": "sales",
        "title": "销售结转成本凭证",
        "eyebrow": "SALES COST CARRY-FORWARD VOUCHER",
        "status": "ready",
        "status_label": "已实现",
        "note": "数据来源：销售台账（结转类型 × 销售大类）。会计科目由平台按结转类型映射，"
                "金额取台账「销售成本」列（结算后按销售数量 × 月末加权平均价回填）。",
        "url_name": "sales_cost_voucher",
        "export_kind": "sales_voucher",
        "metrics": [
            _metric("记账月份", voucher["month"]),
            _metric("合计数量", "{:,.3f}".format(totals["quantity"])),
            _metric("合计金额", "{:,.2f}".format(totals["amount"])),
            _metric("参与汇总台账行数", "{:,}".format(totals["source_rows"])),
            _metric("结转类型", "%d 类" % len(CARRY_COST_TYPES)),
        ],
        "entry_headers": list(voucher["headers"]),
        "entries": entries,
        "balance": {
            "debit": totals["amount"],
            "credit": totals["amount"],
            "balanced": True,
            "label": "金额",
        },
        "checks": checks,
    }


def _purchase_card(period):
    voucher = build_purchase_voucher(period, group="single")
    summary = voucher["summary"]
    entries = []
    for block in voucher["blocks"]:
        for row in block["rows"]:
            # 与采购凭证明细页保持一致的 7 列
            entries.append(
                {
                    "level": row["level"],
                    "cells": [
                        _cell(row["summary"]),
                        _cell(row["account"]),
                        _cell(row["aux"]),
                        _cell("{:,.2f}".format(row["debit"]) if row["debit"] else "", True),
                        _cell("{:,.2f}".format(row["credit"]) if row["credit"] else "", True),
                        _cell(
                            "{:,.3f}".format(row["quantity"]) if row["quantity"] else "",
                            True,
                        ),
                        _cell("{:,}".format(row["source_rows"]) if row["source_rows"] else "", True),
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
        "metrics": [
            _metric("记账月份", voucher["month"]),
            _metric("本月到货金额", "{:,.2f}".format(summary["receipt_amount"])),
            _metric("① 到货到票入账", "{:,.2f}".format(summary["same_month_amount"])),
            _metric("② 到货未到票暂估", "{:,.2f}".format(summary["accrual_amount"])),
            _metric("③ 冲回原暂估", "{:,.2f}".format(summary["reversal_amount"])),
            _metric(
                "本月价差",
                "{:,.2f}".format(summary["difference_total"]),
                tone="danger" if summary["difference_total"] else "",
            ),
            _metric("月末暂估余额", "{:,.2f}".format(summary["ending_balance"])),
        ],
        "entry_headers": ["摘要", "科目", "辅助核算", "借方", "贷方", "数量", "来源行数"],
        "entries": entries,
        "balance": {
            "debit": voucher["totals"]["debit"],
            "credit": voucher["totals"]["credit"],
            "balanced": voucher["totals"]["balanced"],
            "label": "借贷",
        },
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


def build_voucher_center(period):
    """本期凭证中心数据：已实现 2 类 + 规划中 3 类。"""
    # 按建议的生成顺序排列：暂估回冲 → 投料 → 制造费用 → 完工入库 → 销售结转
    cards = [_purchase_card(period)]
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
    cards.append(_sales_card(period))

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
