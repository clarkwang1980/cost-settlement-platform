from __future__ import annotations

import os
import uuid
from collections import defaultdict
from datetime import timedelta
from decimal import Decimal, ROUND_HALF_UP

from django.db import transaction
from django.db.models import Sum
from django.utils import timezone

from settlement.models import (
    HistoricalPrice,
    AccountingPeriod,
    BomLine,
    EndingInventoryBalance,
    ExceptionItem,
    InventoryTransaction,
    InventoryValuation,
    MatchAllocation,
    MatchingMode,
    Material,
    OpeningInventoryBalance,
    OpeningWip,
    ReconciliationResult,
    SettlementRun,
    WorkOrderCostResult,
    WorkshopSettlement,
    WorkshopCostPool,
)
from settlement.services.rollforward import sync_carry_forward
from settlement.services.matching import (
    auto_match_all,
    is_purchase_receipt,
    is_sales_issue,
)
from settlement.services.sales_uninvoiced import (
    build_goods_issued_rollforward,
    build_sales_uninvoiced_rows,
    summarize_current_sales_cost,
    summarize_sales_uninvoiced,
)
from settlement.services.workorder_costing import (
    BULK_UPDATE_BATCH,
    FINISHED_RECEIPT_TYPES,
    calculate_work_order_costs,
    clear_baogong_receipt_prices,
    is_baogong_receipt,
    price_baogong_receipts,
)
from settlement.services.imports import (
    WORKSHOP_ORDER,
    WORKSHOP_LABELS,
    normalize_workshop,
)


ZERO = Decimal("0")
CENT = Decimal("0.01")
PRICE_PRECISION = Decimal("0.00000001")
CONVERGENCE_TOLERANCE = Decimal("0.005")
MAX_ITERATIONS = 50

# 用户 2026-09-23 定：BOM 子项「用量」（分子÷分母）合计 **超过 100%** 时，
# 视作配比类录入偏差 —— 按用量比例缩放到合计 100% 后再算标准成品单价
# （等价于 标准成品单价 = Σ(用量×单价) ÷ Σ用量）。
# 合计 ≤ 100% 的工单不缩放（只处理"超过"的情形，用户明确要求）。
# 置 0 或 None 可关闭该规则（回到 Σ(用量×单价)）。
BOM_QTY_NORMALIZE_OVER = Decimal("1")


def money(value) -> Decimal:
    return Decimal(value or ZERO).quantize(CENT, rounding=ROUND_HALF_UP)


def price(value) -> Decimal:
    return Decimal(value or ZERO).quantize(
        PRICE_PRECISION, rounding=ROUND_HALF_UP
    )


def clamp_rate(value) -> Decimal:
    value = Decimal(value or ZERO)
    return min(max(value, ZERO), Decimal("1"))


def is_opening_transaction(txn: InventoryTransaction) -> bool:
    return "期初" in txn.business_type


INTERNAL_TRANSFER_TYPES = ("调拨", "库存调整", "库存转换", "报工调整")


def is_internal_transfer(txn: InventoryTransaction) -> bool:
    """内部流转（仓库间移动/调整）：不构成真实的库存增减，不计入库存金额勾稽。
    与导入时忽略匹配的内部流转类型保持一致（调拨、库存调整、库存转换、报工调整）。"""
    bt = txn.business_type or ""
    return any(token in bt for token in INTERNAL_TRANSFER_TYPES)



def _fixed_receipt_amount(txn: InventoryTransaction) -> bool:
    if is_opening_transaction(txn):
        return True
    if is_purchase_receipt(txn):
        return txn.amount is not None and bool(txn.price_source)
    if is_baogong_receipt(txn):
        return txn.amount is not None
    return txn.amount is not None and txn.price_source not in {
        "",
        "月末加权平均",
        "生产入库成本",
    }


def _valuation_pass(period: AccountingPeriod, previous_prices=None):
    previous_prices = previous_prices or {}
    grouped = defaultdict(list)
    for txn in period.inventory_transactions.select_related("material").all():
        grouped[txn.material_code].append(txn)
    # 期初可能按仓库/批次多行：按物料汇总数量金额
    automatic_openings = defaultdict(
        lambda: {"quantity": ZERO, "amount": ZERO, "material_name": ""}
    )
    for item in OpeningInventoryBalance.objects.filter(period=period):
        row = automatic_openings[item.material_code]
        row["quantity"] += Decimal(item.quantity or ZERO)
        row["amount"] += Decimal(item.amount or ZERO)
        if item.material_name and not row["material_name"]:
            row["material_name"] = item.material_name
    for material_code in automatic_openings:
        grouped.setdefault(material_code, [])

    summaries = {}
    prices = {}
    for material_code, txns in grouped.items():
        opening_qty = ZERO
        opening_amount = ZERO
        has_manual_opening = False
        receipt_qty = ZERO
        # 参与**月末加权价**计算的入库数量：只含正常匹配的入库（不含内部流转，
        # 内部流转虽然算期末数量，但它们的金额不计成本，不能进单价分母 —— 否则单价会被摊薄）
        priced_receipt_qty = ZERO
        fixed_receipt_amount = ZERO
        floating_receipts = []
        issue_qty = ZERO

        for txn in txns:
            qty = abs(txn.quantity)
            # 被忽略匹配的流水（**内部流转**：调拨/库存调整/库存转换/报工调整）：
            #   **数量照算**（它们也是真实的数量移动），只是不参与金额与月末加权价计算。
            #   用户 2026-09-30 实测 14.298：库存调整 +3.01（入库）原来被整行跳过 →
            #   期末 869.53 ＋ 0 − 872.54 ＝ −3.01 ✗；加上调整后应为 0.00 ✓。
            if txn.match_mode == MatchingMode.IGNORED:
                if is_opening_transaction(txn):
                    has_manual_opening = True
                    opening_qty += qty
                    opening_amount += Decimal(txn.amount or ZERO)
                elif txn.direction == InventoryTransaction.Direction.IN:
                    receipt_qty += qty
                else:
                    issue_qty += qty
                continue
            qty = abs(txn.quantity)
            if is_opening_transaction(txn):
                has_manual_opening = True
                opening_qty += qty
                opening_amount += Decimal(txn.amount or ZERO)
            elif txn.direction == InventoryTransaction.Direction.IN:
                receipt_qty += qty
                priced_receipt_qty += qty
                if _fixed_receipt_amount(txn):
                    fixed_receipt_amount += Decimal(txn.amount or ZERO)
                else:
                    floating_receipts.append(txn)
            else:
                issue_qty += qty
        automatic_opening = automatic_openings.get(material_code)
        if automatic_opening and not has_manual_opening:
            opening_qty = Decimal(automatic_opening["quantity"] or ZERO)
            opening_amount = Decimal(automatic_opening["amount"] or ZERO)

        prior_price = Decimal(previous_prices.get(material_code, ZERO))
        floating_qty = sum((abs(txn.quantity) for txn in floating_receipts), ZERO)
        numerator = opening_amount + fixed_receipt_amount + floating_qty * prior_price
        denominator = opening_qty + priced_receipt_qty
        weighted_price = numerator / denominator if denominator else ZERO

        # Returns and other unpriced receipts are valued at the same monthly average.
        # One additional algebraic step removes their self-referential effect.
        if floating_qty and denominator > floating_qty:
            weighted_price = (
                opening_amount + fixed_receipt_amount
            ) / (denominator - floating_qty)
        weighted_price = price(weighted_price)
        prices[material_code] = weighted_price

        for txn in txns:
            if txn.match_mode == MatchingMode.IGNORED:
                continue
            qty = abs(txn.quantity)
            if txn.direction == InventoryTransaction.Direction.OUT:
                txn.unit_price = weighted_price
                txn.amount = money(qty * weighted_price)
                txn.price_source = "月末加权平均"
                txn.save(update_fields=["unit_price", "amount", "price_source"])
            elif (
                not is_opening_transaction(txn)
                and not _fixed_receipt_amount(txn)
                and not is_baogong_receipt(txn)
            ):
                txn.unit_price = weighted_price
                txn.amount = money(qty * weighted_price)
                txn.price_source = "月末加权平均"
                txn.save(update_fields=["unit_price", "amount", "price_source"])

        total_receipt_amount = sum(
            (
                Decimal(txn.amount or ZERO)
                for txn in txns
                if txn.match_mode != MatchingMode.IGNORED
                and txn.direction == InventoryTransaction.Direction.IN
                and not is_opening_transaction(txn)
            ),
            ZERO,
        )
        total_issue_amount = sum(
            (
                Decimal(txn.amount or ZERO)
                for txn in txns
                if txn.match_mode != MatchingMode.IGNORED
                and txn.direction == InventoryTransaction.Direction.OUT
            ),
            ZERO,
        )
        ending_qty = opening_qty + receipt_qty - issue_qty
        ending_amount = money(ending_qty * weighted_price)
        material_name = next(
            (txn.material_name for txn in txns if txn.material_name),
            automatic_opening["material_name"] if automatic_opening else "",
        )
        summaries[material_code] = {
            "material_name": material_name,
            "opening_quantity": opening_qty,
            "opening_amount": money(opening_amount),
            "receipt_quantity": receipt_qty,
            "receipt_amount": money(total_receipt_amount),
            "issue_quantity": issue_qty,
            "issue_amount": money(total_issue_amount),
            "weighted_average_price": weighted_price,
            "ending_quantity": ending_qty,
            "ending_amount": ending_amount,
        }
    return prices, summaries


def _persist_valuations(period, summaries, iterations, converged):
    InventoryValuation.objects.filter(period=period).delete()
    InventoryValuation.objects.bulk_create(
        [            InventoryValuation(
                period=period,
                material_code=material_code,
                material_name=data["material_name"],
                opening_quantity=data["opening_quantity"],
                opening_amount=data["opening_amount"],
                receipt_quantity=data["receipt_quantity"],
                receipt_amount=data["receipt_amount"],
                issue_quantity=data["issue_quantity"],
                issue_amount=data["issue_amount"],
                weighted_average_price=data["weighted_average_price"],
                ending_quantity=data["ending_quantity"],
                ending_amount=data["ending_amount"],
                iterations=iterations,
                converged=converged,
            )
            for material_code, data in summaries.items()
        ],
        batch_size=BULK_UPDATE_BATCH,
    )
    # 期末存货以用户上传报表为准（数量不变），单价金额由系统按月末加权价回填
    _fill_ending_prices(period, summaries)
    # 系统自动期末：对"无人工期末"的物料按 期初+入库−出库(FIFO批次) 生成
    # （浆粕成品等未上传期末的物料由此自动算出期末数；有人工期末的物料以人工为准）
    _persist_ending_by_batch(period, summaries)


def _fill_ending_prices(period, summaries):
    """用户上传的期末存货：数量保持不变，回填单价金额 = 月末加权价 × 数量。"""
    for ending in period.ending_inventory_balances.filter(
        is_system_generated=False
    ):
        if ending.material_code not in summaries:
            continue
        price = summaries[ending.material_code]["weighted_average_price"]
        ending.unit_price = price
        ending.amount = money(Decimal(ending.quantity or ZERO) * price)
        ending.save(update_fields=["unit_price", "amount"])




def _persist_ending_by_batch(period, summaries):
    """按 物料+仓库+批次 生成期末存货（批次级）。

    规则：
    1. 期初各仓库/批次行 = 期初库存，顺序按期初行（即批次入库先后）
    2. 本月入库流水：能匹配到同仓库同批次的，累加到该批次；
       匹配不到或没有批次号的，新建一行（新批次/无批次）
    3. 本月出库流水：按先进先出（FIFO）从最早批次扣减
    4. 期末行 = 剩余数量非零的批次行；金额 = 数量 × 物料月末加权价
    """
    EndingInventoryBalance.objects.filter(
        period=period, is_system_generated=True
    ).delete()
    manual_codes = set(
        EndingInventoryBalance.objects.filter(
            period=period, is_system_generated=False
        ).values_list("material_code", flat=True)
    )
    month_end = (
        (period.month.replace(day=28) + timedelta(days=4)).replace(day=1)
        - timedelta(days=1)
    )

    # 期初批次行（按 id 顺序 = 上传顺序 = 先进先出顺序）
    opening_rows = {}
    for opening in OpeningInventoryBalance.objects.filter(period=period).order_by(
        "id"
    ):
        code = opening.material_code
        key = (opening.warehouse or "", opening.batch_no or "")
        opening_rows.setdefault(code, []).append(
            {
                "warehouse": opening.warehouse,
                "batch_no": opening.batch_no,
                "category_name": opening.category_name or "",
                "uom": opening.uom or "",
                "qty": Decimal(opening.quantity or ZERO),
            }
        )

    # 本月入库/出库流水（排除被忽略的内部流转）
    # 数量口径按用户上传流水原样（含调拨/库存调整等内部流转，这些也是真实的数量移动）
    txns = list(
        period.inventory_transactions.all().order_by("txn_date", "id")
    )

    # 物料 → 分类（用户 2026-09-19：期末存货行以前**不带分类**，导出里「分类名称」列是空的，
    # 「库存金额结转」只能靠期初/流水反查分类；承接给下期时还会把空分类带过去）。
    # 这里按 期初存货 → 本月流水 的顺序取第一个非空分类，写进期末行。
    category_of = {}
    for opening in OpeningInventoryBalance.objects.filter(period=period).order_by("id"):
        if opening.material_code and (opening.category_name or "").strip():
            category_of.setdefault(opening.material_code, opening.category_name.strip())
    for txn in txns:
        if txn.material_code and (txn.category_name or "").strip():
            category_of.setdefault(txn.material_code, txn.category_name.strip())

    # 期初类型流水（业务类型含"期初"的入库）也作为初始批次行
    for txn in txns:
        if (
            not is_opening_transaction(txn)
            or txn.direction != InventoryTransaction.Direction.IN
        ):
            continue
        code = txn.material_code
        key = (txn.warehouse or "", txn.batch_no or "")
        rows = opening_rows.setdefault(code, [])
        matched = None
        for row in rows:
            if (
                (row["warehouse"] or "") == key[0]
                and (row["batch_no"] or "") == key[1]
            ):
                matched = row
                break
        if matched:
            matched["qty"] += Decimal(txn.quantity or ZERO)
        else:
            rows.append(
                {
                    "warehouse": txn.warehouse or None,
                    "batch_no": txn.batch_no or None,
                    "category_name": txn.category_name or "",
                    "uom": txn.uom or "",
                    "qty": Decimal(txn.quantity or ZERO),
                }
            )

    ending_by_code = {}
    codes = set(opening_rows) | {
        t.material_code for t in txns
    }
    # 用户 2026-09-19：**浆粕（浆粕/微纤/JP-）也一并结转** ——
    # 以前这里把浆粕成品整类排除在通用期末表之外（"浆粕成品期末独立核算"），结果是：
    # 浆粕没有期末存货行 → 下期承接（_carry_inventory 优先读上期期末存货）拿不到它 → 库存断了。
    # 现在浆粕与普通物料同规则：期末数量 = 期初（浆粕期初上传）+ 本月入库 − 本月出库（浆粕流水上传），
    # 金额 = 数量 × 月末加权价（来自浆粕成本核算）。
    # 若用户另外上传了浆粕期末，则走上面的 manual_codes 分支、以人工为准（保持原有优先级）。
    for code in sorted(codes):
        if code in manual_codes:
            continue
        # 期末数量按物料汇总：期初（各批号合计）+ 本月全部入库 − 本月全部出库
        # （含调拨/库存转换等内部流转，成对抵消）。不做批号 FIFO 扣减，避免调拨顺序/批号错位产生假负数。
        # 不做批号 FIFO 扣减，避免调拨顺序/批号错位产生假负数。
        opening_qty = sum(
            (row["qty"] for row in opening_rows.get(code, [])),
            ZERO,
        )
        # 用户 2026-10-17：**期初就该承接上月期末**。没有上传期初的物料（浆粕等），
        #   直接取上期期末，避免"上月有结存、本月期初为 0"。
        if opening_qty == ZERO:
            _prev = (
                AccountingPeriod.objects.filter(month__lt=period.month)
                .order_by("-month")
                .first()
            )
            if _prev is not None:
                _pv = _prev.inventory_valuations.filter(material_code=code).first()
                if _pv is not None:
                    _end = Decimal(_pv.ending_quantity or 0)
                    if _end != ZERO:
                        opening_qty = _end
                        opening_rows.setdefault(code, []).append(
                            {"qty": _end, "amount": Decimal(_pv.ending_amount or 0)}
                        )
        in_qty = ZERO
        out_qty = ZERO
        for txn in txns:
            if txn.material_code != code or is_opening_transaction(txn):
                continue
            q = abs(Decimal(txn.quantity or ZERO))
            if txn.direction == InventoryTransaction.Direction.IN:
                in_qty += q
            else:
                out_qty += q
        end_qty = opening_qty + in_qty - out_qty
        if end_qty == ZERO:
            continue
        weighted_price = ZERO
        if code in summaries:
            weighted_price = summaries[code]["weighted_average_price"]
        ending_by_code[code] = [
            {
                "material_name": summaries[code]["material_name"] if code in summaries else "",
                "category_name": category_of.get(code, ""),
                "warehouse": None,
                "batch_no": None,
                "uom": "",
                "qty": end_qty,
                "amount": money(end_qty * weighted_price),
                "unit_price": weighted_price,
            }
        ]

    rows_to_create = []
    for code, rows in ending_by_code.items():
        for row in rows:
            rows_to_create.append(
                EndingInventoryBalance(
                    period=period,
                    material_code=code,
                    material_name=row["material_name"],
                    ending_date=month_end,
                    category_name=row["category_name"] or category_of.get(code, ""),
                    warehouse=row["warehouse"],
                    batch_no=row["batch_no"],
                    uom=row["uom"],
                    quantity=row["qty"],
                    amount=row["amount"],
                    unit_price=row["unit_price"],
                    is_system_generated=True,
                )
            )
    EndingInventoryBalance.objects.bulk_create(
        rows_to_create, batch_size=BULK_UPDATE_BATCH
    )


def _add_exception(period, category, severity, message, reference=""):
    ExceptionItem.objects.create(
        period=period,
        category=category,
        severity=severity,
        message=message,
        reference=reference,
    )


def _build_exceptions(period, converged):
    period.exceptions.all().delete()
    if not converged:
        _add_exception(
            period,
            "循环计算",
            ExceptionItem.Severity.BLOCKING,
            f"生产入库成本在{MAX_ITERATIONS}次循环后仍未收敛。",
        )
    for txn in period.inventory_transactions.all():
        if (
            txn.match_mode != MatchingMode.IGNORED
            and is_purchase_receipt(txn)
            and (
                txn.amount is None
                or txn.unit_price is None
                or txn.unit_price == 0
            )
        ):
            _add_exception(
                period,
                "采购入库",
                ExceptionItem.Severity.BLOCKING,
                "采购入库未匹配合同价格或发票价格。",
                f"{txn.document_no}/{txn.material_code}",
            )
        if (
            txn.match_mode != MatchingMode.IGNORED
            and txn.direction == InventoryTransaction.Direction.OUT
            and txn.quantity
            and Decimal(txn.amount or ZERO) == 0
        ):
            _add_exception(
                period,
                "库存计价",
                ExceptionItem.Severity.BLOCKING,
                "出库数量不为零但出库金额为零。",
                f"{txn.document_no}/{txn.material_code}",
            )
    for invoice in period.purchase_invoices.filter(
        remaining_quantity__gt=0
    ).exclude(match_mode=MatchingMode.IGNORED):
        _add_exception(
            period,
            "采购发票",
            ExceptionItem.Severity.WARNING,
            f"采购发票仍有{invoice.remaining_quantity}未匹配入库。",
            invoice.invoice_no,
        )
    for invoice in period.sales_invoices.filter(
        remaining_quantity__gt=0
    ).exclude(match_mode=MatchingMode.IGNORED):
        _add_exception(
            period,
            "销售发票",
            ExceptionItem.Severity.WARNING,
            f"销售发票仍有{invoice.remaining_quantity}未匹配销售出库。",
            invoice.invoice_no,
        )
    uninvoiced_rows = build_sales_uninvoiced_rows(period)
    if uninvoiced_rows:
        uninvoiced = summarize_sales_uninvoiced(
            uninvoiced_rows,
            period,
        )
        _add_exception(
            period,
            "已出库未开票",
            ExceptionItem.Severity.INFO,
            (
                f"月末尚有{uninvoiced['issue_count']}笔销售出库未全部开票，"
                f"未开票数量{uninvoiced['outstanding_quantity']}，"
                f"发出商品成本{money(uninvoiced['outstanding_cost'])}。"
            ),
            f"{period.month:%Y-%m}",
        )
    for result in period.work_order_cost_results.select_related("work_order").all():
        order = result.work_order
        if not result.direct_material and not period.opening_wip.filter(
            order_no=order.order_no
        ).exists():
            _add_exception(
                period,
                "工单成本",
                ExceptionItem.Severity.WARNING,
                "生产工单没有材料成本。",
                order.order_no,
            )
        if (
            result.finished_quantity
            and not result.wip_quantity
            and not order.plan_quantity
        ):
            _add_exception(
                period,
                "工单成本",
                ExceptionItem.Severity.WARNING,
                "工单未填计划数量，本期按全部完工处理，在产品为0。",
                order.order_no,
            )
        if not period.workshop_cost_pools.filter(
            workshop=order.production_dept_code
        ).exists():
            _add_exception(
                period,
                "车间费用",
                ExceptionItem.Severity.WARNING,
                "该车间没有导入本月费用池，制造费用按零计算。",
                order.production_dept_code,
            )
    # 投料物料无月末加权价
    priced_codes = set(
        period.inventory_valuations.values_list("material_code", flat=True)
    )
    for code in (
        period.material_issues.exclude(material_code="")
        .values_list("material_code", flat=True)
        .distinct()
    ):
        valuation = (
            period.inventory_valuations.filter(material_code=code).first()
        )
        if code not in priced_codes or (
            valuation is not None and valuation.weighted_average_price == 0
        ):
            _add_exception(
                period,
                "投料计价",
                ExceptionItem.Severity.WARNING,
                "投料物料没有月末加权平均价，材料成本按0计算。",
                code,
            )
    # 报工入库未按完工成本计价
    for txn in period.inventory_transactions.filter(
        business_type__in=FINISHED_RECEIPT_TYPES
    ).exclude(match_mode=MatchingMode.IGNORED):
        if txn.amount is None:
            _add_exception(
                period,
                "完工入库",
                ExceptionItem.Severity.WARNING,
                "报工入库未按工单完工成本计价（工单无完工成本）。",
                f"{txn.document_no}/{txn.material_code}",
            )
    # 报工记录数量与报工入库数量差异
    report_qty = defaultdict(lambda: ZERO)
    for row in (
        period.work_reports.exclude(work_order_no="")
        .values("work_order_no")
        .annotate(total=Sum("quantity"))
    ):
        report_qty[row["work_order_no"]] = Decimal(row["total"] or ZERO)
    receipt_qty = defaultdict(lambda: ZERO)
    for txn in period.inventory_transactions.filter(business_type__in=FINISHED_RECEIPT_TYPES):
        receipt_qty[txn.document_no] += Decimal(txn.quantity or ZERO)
    for order_no, reported in report_qty.items():
        received = receipt_qty.get(order_no, ZERO)
        if abs(reported - received) > Decimal("0.001"):
            _add_exception(
                period,
                "完工入库",
                ExceptionItem.Severity.INFO,
                f"报工数量{reported}与报工入库数量{received}不一致。",
                order_no,
            )
    for valuation in period.inventory_valuations.filter(ending_quantity__lt=0):
        _add_exception(
            period,
            "库存数量",
            ExceptionItem.Severity.WARNING,
            f"期末库存数量为负数：{valuation.ending_quantity}",
            valuation.material_code,
        )


def _reconciliation_status(difference, tolerance=Decimal("0.10")):
    return (
        ReconciliationResult.Status.PASS
        if abs(difference) <= tolerance
        else ReconciliationResult.Status.FAIL
    )


def _build_reconciliations(period):
    # 保留已确认的勾稽（用户确认过的差异，重建后自动恢复，避免重复确认）
    confirmed_codes = set(
        ReconciliationResult.objects.filter(
            period=period, confirmed=True
        ).values_list("check_code", flat=True)
    )
    ReconciliationResult.objects.filter(period=period).delete()

    opening = (
        period.inventory_valuations.aggregate(total=Sum("opening_amount"))["total"]
        or ZERO
    )
    receipts = (
        period.inventory_valuations.aggregate(total=Sum("receipt_amount"))["total"]
        or ZERO
    )
    issues = (
        period.inventory_valuations.aggregate(total=Sum("issue_amount"))["total"]
        or ZERO
    )
    # 期末金额：一律用**平台计算的期末**（期初+入库−出库）。
    # 用户 2026-10-13：期末数是平台自己算出来的，不应以上传的期末存货表为准
    #   （按上传表会把"ERP 期末"当账面，等式两边口径不同，必然出差异）。
    #   上传的期末存货表改成一条**提示**对比：上传合计 vs 平台计算合计。
    ending = money(opening + receipts - issues)
    left = money(opening + receipts)
    right = money(issues + ending)
    difference = money(left - right)

    uploaded_ending = period.ending_inventory_balances.filter(
        is_system_generated=False
    )
    if uploaded_ending.exists():
        uploaded_total = (
            uploaded_ending.aggregate(total=Sum("amount"))["total"] or ZERO
        )
        gap = money(uploaded_total - ending)
        if abs(gap) > Decimal("0.10"):
            # 逐物料核对：列出差额最大的几个物料，方便对账（用户 2026-10-14）
            uploaded_map = {
                b.material_code: Decimal(b.amount or ZERO) for b in uploaded_ending
            }
            calc_map = {
                v.material_code: Decimal(v.ending_amount or ZERO)
                for v in period.inventory_valuations.all()
            }
            diffs = []
            for code in set(uploaded_map) | set(calc_map):
                up = uploaded_map.get(code, ZERO)
                calc = calc_map.get(code, ZERO)
                delta = up - calc
                if abs(delta) > Decimal("1"):
                    diffs.append((abs(delta), code, up, calc, delta))
            diffs.sort(reverse=True)
            # 用户 2026-10-16：提示分两类，避免每月都把"未上传期末的物料（如浆粕车间）"
            #   当成对不上的差异。
            auto_only = [d for d in diffs if d[1] not in uploaded_map]   # 平台有、上传没有
            both = [d for d in diffs if d[1] in uploaded_map]            # 两边都有但金额不同
            auto_total = sum((d[4] for d in auto_only), ZERO)
            both_total = sum((d[4] for d in both), ZERO)
            detail = "；".join(
                f"{code} 上传{up:,.2f}／平台{calc:,.2f}（差{delta:+,.2f}）"
                for _a, code, up, calc, delta in both[:5]
            )
            parts = [
                f"上传的期末存货与平台计算期末差 {gap:,.2f}：",
                f"①【你上传里没有、平台按「期初＋入库−出库」自动算的物料】"
                f"{len(auto_only)} 个、合计 {abs(auto_total):,.2f}"
                f"（例：未上传期末的浆粕车间，属正常，不影响锁账）",
                f"②【两边都有但对不上的物料】{len(both)} 个、合计 {abs(both_total):,.2f}",
            ]
            if detail:
                parts.append(f"差额最大的：{detail}")
            parts.append("（仅提示，不影响锁账）")
            ExceptionItem.objects.create(
                period=period,
                category="库存金额",
                severity=ExceptionItem.Severity.INFO,
                message="".join(parts),
                reference="期末存货表",
            )
    ReconciliationResult.objects.create(
        period=period,
        check_code="inventory_amount",
        check_name="库存金额：期初＋入库＝出库＋期末",
        left_amount=left,
        right_amount=right,
        difference=difference,
        status=_reconciliation_status(difference),
    )

    material_flow = ZERO
    for txn in period.inventory_transactions.all():
        if (txn.business_type or "").strip() not in {
            "投料",
            "投料回撤",
            "投料撤回",
        }:
            continue
        signed = Decimal(txn.amount or ZERO)
        material_flow += (
            signed
            if txn.direction == InventoryTransaction.Direction.OUT
            else -signed
        )
    opening_material = (
        period.opening_wip.aggregate(total=Sum("material"))["total"] or ZERO
    )
    order_current_material = (
        period.work_order_cost_results.aggregate(total=Sum("direct_material"))[
            "total"
        ]
        or ZERO
    ) - opening_material
    left = money(material_flow)
    right = money(order_current_material)
    difference = money(left - right)
    ReconciliationResult.objects.create(
        period=period,
        check_code="work_order_material",
        check_name="投料出库金额＝工单本期材料成本",
        left_amount=left,
        right_amount=right,
        difference=difference,
        status=_reconciliation_status(difference),
    )

    opening_total = sum((item.total for item in period.opening_wip.all()), ZERO)
    pool_total = ZERO
    for pool in period.workshop_cost_pools.all():
        # 与工单成本的分摊口径一致：7 项制造费用全部计入本期投入（含染料助剂、加工费）
        pool_total += (
            pool.water_power
            + pool.steam
            + pool.natural_gas
            + pool.depreciation
            + pool.consumables
            + pool.dye_auxiliary
            + pool.processing_fee
            + pool.wages
        )
    finished_total = (
        period.work_order_cost_results.aggregate(total=Sum("finished_cost"))[
            "total"
        ]
        or ZERO
    )
    wip_total = (
        period.work_order_cost_results.aggregate(total=Sum("wip_cost"))["total"]
        or ZERO
    )
    left = money(opening_total + material_flow + pool_total)
    right = money(finished_total + wip_total)
    difference = money(left - right)
    ReconciliationResult.objects.create(
        period=period,
        check_code="work_order_cost",
        check_name="期初在制＋本期投入＝完工成本＋期末在产品",
        left_amount=left,
        right_amount=right,
        difference=difference,
        status=_reconciliation_status(difference),
    )

    baogong_receipt_amount = sum(
        (
            Decimal(txn.amount or ZERO)
            for txn in period.inventory_transactions.all()
            if (txn.business_type or "").strip() in FINISHED_RECEIPT_TYPES
        ),
        ZERO,
    )
    left = money(finished_total)
    right = money(baogong_receipt_amount)
    difference = money(left - right)
    ReconciliationResult.objects.create(
        period=period,
        check_code="work_order_receipt",
        check_name="工单完工成本＝报工入库金额",
        left_amount=left,
        right_amount=right,
        difference=difference,
        status=_reconciliation_status(difference),
    )

    sales = summarize_current_sales_cost(period)
    right = money(
        sales["invoiced_cost"]
        + sales["goods_issued"]
        + sales["sample_cost"]
        + sales["ignored_cost"]
    )
    difference = money(sales["sales_cost"] - right)
    ReconciliationResult.objects.create(
        period=period,
        check_code="sales_cost",
        check_name="本期销售台账成本＝已开票结转＋未开票发出商品＋样品",
        left_amount=sales["sales_cost"],
        right_amount=right,
        difference=difference,
        status=_reconciliation_status(difference),
        details={
            "invoiced_cost": str(sales["invoiced_cost"]),
            "goods_issued": str(sales["goods_issued"]),
            "sample_cost": str(sales["sample_cost"]),
            "ignored_cost": str(sales["ignored_cost"]),
        },
    )

    goods_issued = build_goods_issued_rollforward(period)
    left = money(
        goods_issued["opening_amount"]
        + goods_issued["current_issue_amount"]
    )
    right = money(
        goods_issued["invoice_cleared_amount"]
        + goods_issued["ending_amount"]
    )
    difference = money(left - right)
    ReconciliationResult.objects.create(
        period=period,
        check_code="goods_issued_rollforward",
        check_name="期初发出商品＋本月出库＝本月开票结转＋期末发出商品",
        left_amount=left,
        right_amount=right,
        difference=difference,
        status=_reconciliation_status(difference),
        details={
            "opening_amount": str(goods_issued["opening_amount"]),
            "current_issue_amount": str(
                goods_issued["current_issue_amount"]
            ),
            "invoice_cleared_amount": str(
                goods_issued["invoice_cleared_amount"]
            ),
            "ending_amount": str(goods_issued["ending_amount"]),
        },
    )


    # 恢复已确认的勾稽
    if confirmed_codes:
        ReconciliationResult.objects.filter(
            period=period, check_code__in=confirmed_codes
        ).update(confirmed=True)


def ensure_workshop_settlements(period: AccountingPeriod):
    workshops = set(
        normalize_workshop(code)
        for code in period.work_orders.exclude(production_dept_code="")
        .values_list("production_dept_code", flat=True)
        if normalize_workshop(code)
    )
    workshops.update(
        normalize_workshop(ws)
        for ws in period.workshop_cost_pools.exclude(workshop="").values_list(
            "workshop", flat=True
        )
        if normalize_workshop(ws)
    )
    existing = {
        item.workshop: item
        for item in period.workshop_settlements.all()
    }
    # 默认结转顺序：WORKSHOP_ORDER（染色车间→纺纱→包覆→针织→机织），
    # 不在列表中的车间按字母序追加到末尾
    ordered = [code for code in WORKSHOP_ORDER if code in workshops]
    ordered += sorted(workshops - set(WORKSHOP_ORDER))
    next_sequence = max(
        (item.sequence for item in existing.values()), default=0
    )
    for workshop in ordered:
        if workshop in existing:
            continue
        next_sequence += 1
        existing[workshop] = WorkshopSettlement.objects.create(
            period=period,
            workshop=workshop,
            sequence=next_sequence,
        )
    stale = set(existing) - workshops
    if stale:
        stale_qs = period.workshop_settlements.filter(
            workshop__in=stale,
            status=WorkshopSettlement.Status.PENDING,
        )
        # 页面每次打开都会调用本函数，因此这里必须避免任何写操作：
        #   1) 没有要清理的车间 → 纯读取，不写库；
        #   2) 正有车间在结转（长事务占着写锁）→ 跳过清理，等结转结束后再清理。
        # 否则页面会等写锁（本地 sqlite 最长 120 秒），表现为页面转圈、打不开。
        settling = SettlementRun.objects.filter(
            period=period,
            status=SettlementRun.Status.RUNNING,
        ).exists()
        if stale_qs.exists() and not settling:
            stale_qs.delete()
    return period.workshop_settlements.order_by("sequence", "workshop")


def update_workshop_sequences(period: AccountingPeriod, sequence_by_id):
    if period.status == AccountingPeriod.Status.LOCKED:
        raise ValueError("该期间已锁账，不能调整车间结转顺序。")
    settlements = list(ensure_workshop_settlements(period))
    if any(
        item.status == WorkshopSettlement.Status.SETTLED
        for item in settlements
    ):
        raise ValueError("已有车间完成结转；如需改序，请先重置车间结转。")
    values = []
    for item in settlements:
        try:
            value = int(sequence_by_id[item.id])
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError(f"{WORKSHOP_LABELS.get(item.workshop, item.workshop)}的结转顺序无效。") from exc
        if value <= 0:
            raise ValueError("结转顺序必须是大于0的整数。")
        values.append(value)
        item.sequence = value
    if len(set(values)) != len(values):
        raise ValueError("不同车间不能使用相同结转顺序。")
    WorkshopSettlement.objects.bulk_update(settlements, ["sequence"])
    return settlements


def reset_workshop_settlements(period: AccountingPeriod):
    if period.status == AccountingPeriod.Status.LOCKED:
        raise ValueError("该期间已锁账，不能重置车间结转。")
    ensure_workshop_settlements(period)
    period.workshop_settlements.update(
        status=WorkshopSettlement.Status.PENDING,
        iterations=0,
        converged=False,
        settled_at=None,
        settled_by=None,
        order_cost=ZERO,
        production_receipt_amount=ZERO,
        error_message="",
    )
    WorkOrderCostResult.objects.filter(period=period).delete()
    InventoryValuation.objects.filter(period=period).delete()
    EndingInventoryBalance.objects.filter(
        period=period, is_system_generated=True
    ).delete()
    ReconciliationResult.objects.filter(period=period).delete()
    period.exceptions.all().delete()
    clear_baogong_receipt_prices(period)
    period.status = AccountingPeriod.Status.OPEN
    period.last_run_at = None
    period.save(update_fields=["status", "last_run_at"])


def _restore_unpriced_receipts(period, before: dict):
    """清空后仍未被重新计价的完工入库行，恢复清空前的计价。

    为什么需要：完工入库行只在"所属车间本次真正结算"时才重新计价；
    已结转的车间（重复结算场景）不会重算，清空后就永久丢金额，
    月末加权平均价会因此偏高/偏低（成本凭空变化）。
    返回恢复的行数。
    """
    from settlement.services.workorder_costing import FINISHED_RECEIPT_TYPES

    rows = []
    for txn in period.inventory_transactions.filter(
        business_type__in=FINISHED_RECEIPT_TYPES
    ):
        if txn.amount not in (None, ZERO):
            continue
        saved = before.get(txn.id)
        if not saved:
            continue
        txn.unit_price, txn.amount, txn.price_source = saved
        rows.append(txn)
    if rows:
        InventoryTransaction.objects.bulk_update(
            rows, ["unit_price", "amount", "price_source"], batch_size=BULK_UPDATE_BATCH
        )
    return len(rows)


def _prepare_period(period):
    carry_summary = sync_carry_forward(period)
    _normalize_priced_production_flows(period)
    auto_match_all(period)
    ensure_workshop_settlements(period)
    # 清空前先记下完工入库行的计价：重复结算时这些行不会再走车间结算，
    # 若清空后没能重新计价，会导致月末加权价失真（成本凭空变化）。
    from settlement.services.workorder_costing import FINISHED_RECEIPT_TYPES

    _before_receipt_prices = {
        t.id: (t.unit_price, t.amount, t.price_source)
        for t in period.inventory_transactions.filter(
            business_type__in=FINISHED_RECEIPT_TYPES
        )
        if t.amount is not None and t.amount != ZERO
    }
    clear_baogong_receipt_prices(period)
    # 清空后用现有工单成本结果立即回填全部报工入库，
    # 保证已结转车间的报工入库始终带完工成本，不被后续单车间结转清空
    price_baogong_receipts(period, None)
    _restore_unpriced_receipts(period, _before_receipt_prices)
    _fill_opening_prices(period)
    return carry_summary


def _fill_opening_prices(period):
    """期初存货只填数量时，用当月采购合同不含税价回填单价、金额。

    金额 = 数量 × 合同不含税价；没有合同价的期初行保持 0（由异常清单提示）。
    """
    contract_prices = {}
    for contract in period.purchase_contracts.order_by(
        "contract_date", "id"
    ):
        contract_prices.setdefault(
            contract.material_code, contract.unit_price_excl_tax
        )
    filled = 0
    for opening in period.opening_inventory_balances.filter(
        is_system_generated=False
    ):
        if (opening.unit_price or ZERO) > 0 or (opening.amount or ZERO) > 0:
            continue
        price = contract_prices.get(opening.material_code)
        if price is None:
            continue
        opening.unit_price = price
        opening.amount = money(Decimal(opening.quantity or ZERO) * price)
        opening.save(update_fields=["unit_price", "amount"])
        filled += 1
    return filled


def _normalize_priced_production_flows(period: AccountingPeriod):
    """历史导入的 投料/报工 流水曾按内部流转"忽略"处理；
    新口径下它们要参与存货计价，结算前自动修正匹配标志与方向。"""
    from settlement.models import InventoryTransaction

    txns = period.inventory_transactions.filter(
        business_type__in=["投料", "投料回撤", "投料撤回", "报工"],
        match_mode=MatchingMode.IGNORED,
    )
    updated = 0
    for txn in txns:
        txn.match_mode = MatchingMode.AUTO
        txn.match_status = InventoryTransaction.MatchStatus.UNMATCHED
        if txn.business_type in {"投料回撤", "投料撤回"}:
            txn.direction = InventoryTransaction.Direction.IN
        txn.save(
            update_fields=["match_mode", "match_status", "direction"]
            if txn.business_type in {"投料回撤", "投料撤回"}
            else ["match_mode", "match_status"]
        )
        updated += 1
    return updated


def _settle_workshop_core(period, settlement, user=None):
    # 粗纱链处理已改为计算层（build_roving_plan 内置于 calculate_work_order_costs），
    # 不再自动生成任何库存流水/投料记录，原始数据保持干净。
    # 浆粕车间：特殊车间（无工单/无投料/无报工），成本归集 = 原料领料(用户上传浆粕流水) + 车间费用，
    # 按产量分摊到 4 种成品，回填生产入库单价（见 pulp_costing）
    if settlement.workshop == "浆粕车间":
        from settlement.services.pulp_costing import (
            collect_pulp_costs,
            price_pulp_receipts,
        )
        costs = collect_pulp_costs(period)
        max_change = price_pulp_receipts(period)
        converged = max_change <= CONVERGENCE_TOLERANCE
        iteration = 1
        for _ in range(2, 4):
            max_change = price_pulp_receipts(period)
            iteration += 1
            converged = max_change <= CONVERGENCE_TOLERANCE
            if max_change <= CONVERGENCE_TOLERANCE:
                break
        settlement.status = (
            WorkshopSettlement.Status.SETTLED
            if converged
            else WorkshopSettlement.Status.FAILED
        )
        settlement.iterations = iteration
        settlement.converged = converged
        settlement.settled_at = timezone.now()
        settlement.settled_by = user
        settlement.order_cost = money(sum(
            (Decimal(c["total"]) for c in costs["products"].values()), ZERO
        ))
        settlement.production_receipt_amount = money(sum(
            (
                Decimal(t.amount or ZERO)
                for t in period.inventory_transactions.filter(business_type="生产入库")
                if (t.material_name or "") and (
                    "浆粕" in t.material_name or "微纤" in t.material_name
                    or str(t.material_code or "").upper().startswith("JP-")
                )
            ),
            ZERO,
        ))
        settlement.error_message = (
            "" if converged else f"{MAX_ITERATIONS}次循环后仍未收敛。"
        )
        settlement.save(
            update_fields=[
                "status",
                "iterations",
                "converged",
                "settled_at",
                "settled_by",
                "order_cost",
                "production_receipt_amount",
                "error_message",
            ]
        )
        return {
            "iterations": iteration,
            "converged": converged,
            "workshop": settlement.workshop,
            "summary": {"workshop": settlement.workshop, "all_settled": False},
        }

    previous_prices = {}
    summaries = {}
    converged = False
    iteration = 0
    workshop_filter = [settlement.workshop]
    for iteration in range(1, MAX_ITERATIONS + 1):
        previous_prices, summaries = _valuation_pass(
            period, previous_prices
        )
        calculate_work_order_costs(period, workshop_filter, previous_prices)
        max_change = price_baogong_receipts(
            period, workshop_filter
        )
        if max_change <= CONVERGENCE_TOLERANCE:
            converged = True
            break

    previous_prices, summaries = _valuation_pass(period, previous_prices)
    calculate_work_order_costs(period, workshop_filter, previous_prices)
    price_baogong_receipts(period, workshop_filter)
    previous_prices, summaries = _valuation_pass(period, previous_prices)

    settlement.status = (
        WorkshopSettlement.Status.SETTLED
        if converged
        else WorkshopSettlement.Status.FAILED
    )
    settlement.iterations = iteration
    settlement.converged = converged
    settlement.settled_at = timezone.now()
    settlement.settled_by = user
    settlement.order_cost = money(
        period.work_order_cost_results.filter(
            work_order__production_dept_code=settlement.workshop
        ).aggregate(total=Sum("total_cost"))["total"]
        or ZERO
    )
    settlement.production_receipt_amount = money(
        period.inventory_transactions.filter(
            business_type__in=FINISHED_RECEIPT_TYPES,
            document_no__in=[
                item.work_order.order_no
                for item in period.work_order_cost_results.select_related(
                    "work_order"
                ).filter(
                    work_order__production_dept_code=settlement.workshop
                )
            ],
        ).aggregate(total=Sum("amount"))["total"]
        or ZERO
    )
    settlement.error_message = (
        "" if converged else f"{MAX_ITERATIONS}次循环后仍未收敛。"
    )
    settlement.save(
        update_fields=[
            "status",
            "iterations",
            "converged",
            "settled_at",
            "settled_by",
            "order_cost",
            "production_receipt_amount",
            "error_message",
        ]
    )
    if not converged:
        raise ValueError(
            f"{WORKSHOP_LABELS.get(settlement.workshop, settlement.workshop)}完工入库成本在"
            f"{MAX_ITERATIONS}次循环后仍未收敛。"
        )
    return {
        "iterations": iteration,
        "summaries": summaries,
    }


def _global_cost_finalize(period: AccountingPeriod):
    """全部车间结转后，用最终月末加权价重算全部工单成本与报工入库，
    确保工单成本结果与库存流水按同一价格口径，勾稽才能闭合。

    返回 (最终价格, 估值汇总)。
    """
    previous_prices = {}
    summaries = {}
    converged = False
    for _ in range(MAX_ITERATIONS):
        previous_prices, summaries = _valuation_pass(
            period, previous_prices
        )
        calculate_work_order_costs(period, None, previous_prices)
        max_change = price_baogong_receipts(period, None)
        if max_change <= CONVERGENCE_TOLERANCE:
            converged = True
            break
    # 终值对齐：以最终价格重算工单成本（与最终流水金额同价）
    previous_prices, summaries = _valuation_pass(
        period, previous_prices
    )
    calculate_work_order_costs(period, None, previous_prices)
    price_baogong_receipts(period, None)
    previous_prices, summaries = _valuation_pass(
        period, previous_prices
    )
    calculate_work_order_costs(period, None, previous_prices)

    # 同步车间结转汇总数字
    for settlement in period.workshop_settlements.all():
        if settlement.workshop == "浆粕车间":
            continue  # 浆粕无工单：成本已在浆粕分支回填，不覆盖
        settlement.order_cost = money(
            period.work_order_cost_results.filter(
                work_order__production_dept_code=settlement.workshop
            ).aggregate(total=Sum("total_cost"))["total"]
            or ZERO
        )
        settlement.production_receipt_amount = money(
            period.inventory_transactions.filter(
                business_type__in=FINISHED_RECEIPT_TYPES,
                document_no__in=[
                    item.work_order.order_no
                    for item in period.work_order_cost_results.select_related(
                        "work_order"
                    ).filter(
                        work_order__production_dept_code=settlement.workshop
                    )
                ],
            ).aggregate(total=Sum("amount"))["total"]
            or ZERO
        )
        settlement.save(
            update_fields=["order_cost", "production_receipt_amount"]
        )
    return previous_prices, summaries, converged


def _finalize_period_state(period):
    settlements = list(period.workshop_settlements.all())
    all_settled = all(
        item.status == WorkshopSettlement.Status.SETTLED
        and item.converged
        for item in settlements
    )
    total_iterations = sum((item.iterations for item in settlements), 0)
    if all_settled:
        _prices, summaries, _converged = _global_cost_finalize(period)
    else:
        _prices, summaries = _valuation_pass(period, {})
    _persist_valuations(
        period,
        summaries,
        total_iterations,
        all_settled,
    )
    if all_settled:
        _build_exceptions(period, True)
        _build_reconciliations(period)
        period.status = AccountingPeriod.Status.CALCULATED
        period.last_run_at = timezone.now()
        period.save(update_fields=["status", "last_run_at"])
    else:
        ReconciliationResult.objects.filter(period=period).delete()
        period.exceptions.all().delete()
        period.status = AccountingPeriod.Status.OPEN
        period.last_run_at = timezone.now()
        period.save(update_fields=["status", "last_run_at"])
    return {
        "all_settled": all_settled,
        "total_iterations": total_iterations,
    }


# 结转并发控制：以数据库里的 RUNNING 批次为准（跨进程有效）。
#
# 以前用的是进程内 threading.Lock，有两个真问题：
#   1) 线上 gunicorn 默认 3 个 worker，锁不跨进程 —— 两个车间可能同时结转，
#      成本结果互相覆盖；
#   2) 进程被强杀（gunicorn 超时 kill、关窗口）时锁不会释放 —— 界面会永远提示
#      “上一次结转还在进行中”，自己好不了，只能重启服务。
#
# 现在的规则：
#   * 有 RUNNING 批次且未超时 → 判定确实有结转在跑，提示里带上车间名和已运行时长；
#   * 有 RUNNING 批次但已超过 SETTLE_STALE_SECONDS → 判定为进程中断留下的僵尸批次，
#     自动标记失败后放行（自愈，不再需要人工干预）。
SETTLE_STALE_SECONDS = int(os.getenv("SETTLE_STALE_SECONDS", "900"))


class SettlePreconditionError(ValueError):
    """结转的前置条件不满足 —— 属于用户操作层面的提示，不是系统故障。

    用于「这个车间本期已经结转完成」「前序车间还没结转」这类情况：
    应当立刻返回友好提示，既不要白等重算，也不该记成一条"结转失败"。
    """


def _assert_workshop_can_settle(period, settlement):
    """便宜的前置校验（不重算）：已结转 / 前序车间未完成 → 直接抛可读提示。"""
    label = WORKSHOP_LABELS.get(settlement.workshop, settlement.workshop)
    if settlement.status == WorkshopSettlement.Status.SETTLED:
        raise SettlePreconditionError(
            f"{label}本期已经结转完成，无需重复结转。"
            "如确实要重算，请点右上角「重置车间结转」后按顺序重新结转。"
        )
    predecessors = period.workshop_settlements.filter(
        sequence__lt=settlement.sequence
    ).exclude(status=WorkshopSettlement.Status.SETTLED)
    if predecessors.exists():
        names = "、".join(
            WORKSHOP_LABELS.get(code, code)
            for code in predecessors.values_list("workshop", flat=True)
        )
        raise SettlePreconditionError(f"请先完成前序车间：{names}。")


def _claim_settle_run(period, *, mode="workshop", workshop=""):
    """抢占本期结转资格，返回新建的 RUNNING 批次；抢不到就抛中文错误。"""
    with transaction.atomic():
        # 串行化“抢占”这个动作：两个请求同时进来，后到的必然能看到先到的批次。
        AccountingPeriod.objects.select_for_update().get(pk=period.pk)
        active = (
            SettlementRun.objects.filter(
                period=period,
                status=SettlementRun.Status.RUNNING,
            )
            .order_by("-started_at", "-id")
            .first()
        )
        if active is not None:
            started = active.started_at or timezone.now()
            elapsed = int((timezone.now() - started).total_seconds())
            if elapsed < SETTLE_STALE_SECONDS:
                code = (active.summary or {}).get("workshop") or ""
                label = WORKSHOP_LABELS.get(code, code)
                who = f"{label}的结转" if label else "上一次结转"
                raise ValueError(
                    f"{who}还在进行中（已运行 {elapsed} 秒，通常需要 1～2 分钟），"
                    "请等它完成后再操作。"
                )
            active.status = SettlementRun.Status.FAILED
            active.error_message = (
                f"上一次结转运行超过 {SETTLE_STALE_SECONDS // 60} 分钟仍未完成，"
                "判定为服务中断，已自动标记为失败，可以重新结转。"
            )
            active.finished_at = timezone.now()
            active.save(update_fields=["status", "error_message", "finished_at"])
        return SettlementRun.objects.create(
            period=period,
            run_id=str(uuid.uuid4()),
            status=SettlementRun.Status.RUNNING,
            # 抢占时就记下是哪个车间：运行中页面才能显示车间名，
            # 万一中途中断，失败记录也能定位到车间（用户 2026-09-19 反馈）。
            summary={"mode": mode, "workshop": workshop} if workshop else {"mode": mode},
        )


def settle_workshop(period, settlement_id, user=None):
    """车间结转入口：同一时间只允许一个结转在跑（跨进程、可自愈，见 _claim_settle_run）。"""
    workshop = (
        period.workshop_settlements.filter(id=settlement_id)
        .values_list("workshop", flat=True)
        .first()
        or ""
    )
    run = _claim_settle_run(period, mode="workshop", workshop=workshop)
    return _settle_workshop_locked(period, settlement_id, user, run)


def _settle_workshop_locked(period, settlement_id, user=None, run=None):
    if run is None:
        run = _claim_settle_run(period, mode="workshop")
    try:
        with transaction.atomic():
            period = AccountingPeriod.objects.select_for_update().get(pk=period.pk)
            if period.status == AccountingPeriod.Status.LOCKED:
                raise ValueError("该期间已锁账，不能继续结转。")
            # ⚠️ 便宜的校验必须在 _prepare_period（重算约 1～2 分钟）之前：
            # 否则「这个车间已结转 / 前序没完成」也会白等两分钟，最后只弹一句失败
            # （用户 2026-09-19 报的「针织结转不了」就是这个：针织已结转，再点一次白等 2 分钟才报错）。
            settlement = period.workshop_settlements.select_for_update().get(
                id=settlement_id
            )
            _assert_workshop_can_settle(period, settlement)
            carry_summary = _prepare_period(period)
            # 重算可能刷新车间行，取最新状态再校验一次，防止并发下刚被别人结转掉。
            settlement = period.workshop_settlements.select_for_update().get(
                id=settlement_id
            )
            _assert_workshop_can_settle(period, settlement)
            result = _settle_workshop_core(period, settlement, user)
            final_state = _finalize_period_state(period)
            run.status = SettlementRun.Status.SUCCESS
            run.iterations = result["iterations"]
            run.converged = settlement.converged
            run.finished_at = timezone.now()
            run.summary = {
                "mode": "workshop",
                "workshop": settlement.workshop,
                "all_settled": final_state["all_settled"],
                "carry_forward": carry_summary,
            }
            run.save(
                update_fields=[
                    "status",
                    "iterations",
                    "converged",
                    "finished_at",
                    "summary",
                ]
            )
            return run
    except SettlePreconditionError:
        # 用户层面的前置条件不满足（该车间已结转 / 前序未完成）：本次什么都没算，
        # 把抢占到的批次删掉，不留下假的“结转失败”记录（用户 2026-09-19）。
        run.delete()
        raise
    except Exception as exc:
        run.status = SettlementRun.Status.FAILED
        run.error_message = str(exc)
        run.finished_at = timezone.now()
        run.save(update_fields=["status", "error_message", "finished_at"])
        raise


def bom_material_prices(period: AccountingPeriod) -> dict:
    """BOM 子项取价表：物料编码 → 单价。

    口径与投料单价 / BOM 反推**完全一致**（用户 2026-09-30 指出 BOM 表一大堆没单价）：
      ① 该物料本月**月末加权平均价**（库存金额结转，> 0 才算数）；
      ② 没有月末价时用**历史单价**（BOM 子项是被领用的料 → **出库价优先**、入库价兜底）。
    两处都没有的物料不在表里（调用方按 0 处理）。
    """
    prices = {}
    for row in period.inventory_valuations.values(
        "material_code", "weighted_average_price"
    ):
        value = Decimal(row["weighted_average_price"] or ZERO)
        if value > ZERO:
            prices[row["material_code"]] = value
    # 用户 2026-10-15：本月**没进货/没新增**的物料，要**沿用上月（及更早）的月末加权价**。
    #   例：R52.285 在 7 月有月末价、8 月没进货 → 8 月也必须取得到
    #   （否则 BOM 子项缺价、标准成本算不出）。
    previous_values = (
        InventoryValuation.objects.filter(period__month__lt=period.month)
        .order_by("-period__month", "id")
        .values_list("material_code", "weighted_average_price")
    )
    for code, price in previous_values:
        if not code or prices.get(code):
            continue
        value = Decimal(price or ZERO)
        if value > ZERO:
            prices[code] = value
    for item in period.historical_prices.all():
        if prices.get(item.material_code):
            continue
        issue = Decimal(item.issue_price or ZERO)
        receipt = Decimal(item.receipt_price or ZERO)
        if issue > ZERO:
            prices[item.material_code] = issue
        elif receipt > ZERO:
            prices[item.material_code] = receipt
    # 用户 2026-10-14：历史单价**一次上传、后续期间自动沿用**（不要求每月重传）。
    #   本期表里没有的物料，按期间倒序到以前各期的历史单价里找（最近一期优先）。
    previous = (
        HistoricalPrice.objects.filter(period__month__lt=period.month)
        .order_by("-period__month", "id")
    )
    for item in previous:
        code = item.material_code
        if not code or prices.get(code):
            continue
        issue = Decimal(item.issue_price or ZERO)
        receipt = Decimal(item.receipt_price or ZERO)
        if issue > ZERO:
            prices[code] = issue
        elif receipt > ZERO:
            prices[code] = receipt
    return prices


def _fill_bom_prices(period: AccountingPeriod, save: bool = True):
    """回填BOM单价：子项单价 = 子项物料本月月末加权平均价（没有则历史单价兜底，见 bom_material_prices）；
    标准成品单价 = 按工单汇总 Σ(子项用量 × 子项单价)。

    用户 2026-09-23：**子项用量合计超过 100%（Σ用量 > 1）的工单**，先把用量按比例
    缩放到合计 100%，再算标准成品单价（＝ Σ(用量×单价) ÷ Σ用量）；合计 ≤ 100% 的不动。
    用户 2026-09-30：以前只查月末加权价，没结转到的物料一律记 0（BOM 表一大堆没单价）；
    现在与投料单价同口径，补上「历史单价（出库价优先）」兜底。
    save=False 时只算不写库（给 `manage.py fill_bom_prices --dry-run` 用），并返回统计。
    """
    lines = list(period.bom_lines.all())
    if not lines:
        return
    prices = bom_material_prices(period)
    standard_by_order = defaultdict(Decimal)
    qty_by_order = defaultdict(Decimal)
    for line in lines:
        unit_price = Decimal(prices.get(line.sub_item_code, ZERO) or ZERO)
        line.unit_price = unit_price
        standard_by_order[line.work_order_no] += line.qty * unit_price
        qty_by_order[line.work_order_no] += line.qty
    for line in lines:
        total = standard_by_order[line.work_order_no]
        total_qty = qty_by_order[line.work_order_no]
        # 用量合计超过 100% → 缩放到 100%（按用量比例），再算标准成品单价
        if (
            BOM_QTY_NORMALIZE_OVER
            and total_qty > BOM_QTY_NORMALIZE_OVER
        ):
            total = total / total_qty
        line.standard_price = money(total)
    if save:
        BomLine.objects.bulk_update(
            lines, ["unit_price", "standard_price"], batch_size=BULK_UPDATE_BATCH
        )
    return {
        "lines": len(lines),
        "priced": sum(1 for line in lines if (line.unit_price or ZERO) > ZERO),
        "unpriced": sum(1 for line in lines if not (line.unit_price or ZERO) > ZERO),
        "unpriced_codes": len(
            {
                line.sub_item_code
                for line in lines
                if not (line.unit_price or ZERO) > ZERO
            }
        ),
        "orders": len({line.work_order_no for line in lines}),
        "orders_with_standard_price": len(
            {
                line.work_order_no
                for line in lines
                if (line.standard_price or ZERO) > ZERO
            }
        ),
    }


def run_settlement(period: AccountingPeriod, user=None):
    run = _claim_settle_run(period, mode="full")
    try:
        with transaction.atomic():
            period = AccountingPeriod.objects.select_for_update().get(pk=period.pk)
            if period.status == AccountingPeriod.Status.LOCKED:
                raise ValueError("该期间已锁账，不能重新结算。")
            carry_summary = _prepare_period(period)
            total_iterations = 0
            for settlement in period.workshop_settlements.exclude(
                status=WorkshopSettlement.Status.SETTLED
            ).order_by("sequence", "workshop"):
                result = _settle_workshop_core(period, settlement, user)
                total_iterations += result["iterations"]
            final_state = _finalize_period_state(period)
            converged = final_state["all_settled"]
            total_iterations = final_state["total_iterations"]

            # 回填BOM单价：子项单价=子项物料月末加权价；标准成品单价=Σ(子项用量×子项单价)
            _fill_bom_prices(period)

            # 回填销售台账的 成本单价/成本/毛利/毛利率（台账只给数量与收入，成本按月末加权价算）
            from settlement.services.sales_ledger import fill_sales_ledger_costs

            fill_sales_ledger_costs(period)

            run.status = SettlementRun.Status.SUCCESS
            run.iterations = total_iterations
            run.converged = converged
            run.finished_at = timezone.now()
            run.summary = {
                "inventory_materials": period.inventory_valuations.count(),
                "orders": period.work_order_cost_results.count(),
                "blocking_exceptions": period.exceptions.filter(
                    severity=ExceptionItem.Severity.BLOCKING,
                    status=ExceptionItem.Status.OPEN,
                ).count(),
                "failed_reconciliations": period.reconciliations.filter(
                    status=ReconciliationResult.Status.FAIL
                ).count(),
                "carry_forward": carry_summary,
                "workshop_sequence": list(
                    period.workshop_settlements.order_by(
                        "sequence", "workshop"
                    ).values_list("workshop", flat=True)
                ),
            }
            run.save(
                update_fields=[
                    "status",
                    "iterations",
                    "converged",
                    "finished_at",
                    "summary",
                ]
            )
            period.status = AccountingPeriod.Status.CALCULATED
            period.last_run_at = timezone.now()
            period.save(update_fields=["status", "last_run_at"])
            return run
    except Exception as exc:
        run.status = SettlementRun.Status.FAILED
        run.error_message = str(exc)
        run.finished_at = timezone.now()
        run.save(update_fields=["status", "error_message", "finished_at"])
        raise


def can_lock(period: AccountingPeriod):
    if period.workshop_settlements.exclude(
        status=WorkshopSettlement.Status.SETTLED
    ).exists():
        return False, "仍有车间尚未完成独立结转。"
    latest_run = period.runs.first()
    if not latest_run or latest_run.status != SettlementRun.Status.SUCCESS:
        return False, "尚未完成成功结算。"
    if not latest_run.converged:
        return False, "循环计算尚未收敛。"
    if period.exceptions.filter(
        severity=ExceptionItem.Severity.BLOCKING,
        status=ExceptionItem.Status.OPEN,
    ).exists():
        return False, "仍有未处理的阻断异常。"
    if period.reconciliations.filter(
        status=ReconciliationResult.Status.FAIL,
        confirmed=False,
    ).exists():
        return False, "仍有未通过的勾稽项目（已确认差异的除外）。"
    return True, ""
def recalculate_workshop(period: AccountingPeriod, workshop: str):
    """只重算某个车间的工单成本结果（用户 2026-09-24）。

    用在「投料报工情形速查」保存人工调整之后：数据变了要立刻反映到工单成本，
    但又不能跑整套结算（_prepare_period 要 1~2 分钟）。这里走与车间结转同一套
    估值/计价逻辑，只针对一个车间，秒级返回。

    注意：**不改动车间结转状态**——人工调整后本车间（以及会用到它半成品的下游车间）
    仍需在「车间结转」里重新结转一次，账才闭环；页面会提示。
    """
    workshop = str(workshop or "").strip()
    if not workshop:
        return {"workshop": workshop, "orders": 0}
    previous_prices = {}
    workshop_filter = [workshop]
    for _iteration in range(1, MAX_ITERATIONS + 1):
        previous_prices, _summaries = _valuation_pass(period, previous_prices)
        calculate_work_order_costs(period, workshop_filter, previous_prices)
        max_change = price_baogong_receipts(period, workshop_filter)
        if max_change <= CONVERGENCE_TOLERANCE:
            break
    previous_prices, _summaries = _valuation_pass(period, previous_prices)
    calculate_work_order_costs(period, workshop_filter, previous_prices)
    price_baogong_receipts(period, workshop_filter)
    return {
        "workshop": workshop,
        "orders": period.work_order_cost_results.filter(
            work_order__production_dept_code=workshop
        ).count(),
    }
