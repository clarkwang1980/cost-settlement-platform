# -*- coding: utf-8 -*-
"""销售台账（出库+开票）应用逻辑（v2：样品用“样品发票+手动核销”，不拆流水行）。

- 开票状态=已开票 & 开票期间=出库期间（当月发货当月开票）→ 结转当月销售成本：
    生成系统开票记录（发票号 LEDGER-*，不带订单/客户 → 按物料宽松核销当月流水销售出库行），
    auto_match 核销后 → 结算归入“已开票结转成本”。
- 开票状态=已开票 & 开票期间>出库期间（前期发货本期开票）→ 本期结转（冲前期发出商品）：
    生成带订单/客户的系统开票记录，核销期初发出商品明细(OpeningGoodsIssued)或历史月流水出库行。
- 开票状态=未开票 → 不动作：出库行保持未匹配 → 月末自动挂“发出商品”。
- 开票状态=样品免费 → 生成 LEDGER-SAMPLE 样品开票记录并手动核销当月对应数量出库行 →
    该行被 is_sample_issue 识别为样品，结算计入销售费用(sample_cost)、不进未开票/结转口径。

成本金额：所有出库行（含销售/样品）结算时统一按当月月末加权平均价回填
（settlement.services.costing._valuation_pass），本模块只按数量划分口径，不做金额。
"""
from __future__ import annotations

from collections import defaultdict
from decimal import ROUND_HALF_UP, Decimal

from settlement.models import (
    ImportBatch,
    InventoryTransaction,
    MatchAllocation,
    MatchingMode,
    OpeningGoodsIssued,
    SalesInvoiceLine,
    SalesLedgerLine,
)
from settlement.services.matching import (
    _norm,
    _same_party_strict,
    _txn_matched_quantity,
    auto_match_all,
    is_sales_issue,
    is_sample_issue,
)

ZERO = Decimal("0")


def _fmt(value) -> str:
    return f"{Decimal(value or ZERO):g}"


def _pooled_flow_txns(period):
    """当月流水“销售出库”池：物料编码 -> [行(按日期序)]，仅未忽略、非样品。"""
    pool = defaultdict(list)
    qs = InventoryTransaction.objects.filter(
        period=period, match_mode=MatchingMode.AUTO
    ).select_related("period")
    for txn in qs:
        if is_sales_issue(txn) and not is_sample_issue(txn):
            pool[txn.material_code].append(txn)
    for rows in pool.values():
        rows.sort(key=lambda t: (t.txn_date, t.document_no, t.id))
    return pool


def _sum_qty(rows):
    return sum((abs(Decimal(t.quantity or ZERO)) for t in rows), ZERO)

def _invoiced_row_groups(lines, out_month, code):
    """把台账「当月结转」的已开票行按 订单单号+客户 归组（保持台账行序）。

    台账=你已做好的匹配结果，这里保留订单/客户用于精确落行；无订单号的行退化为按物料顺序核销。
    """
    groups = {}
    order = []
    for line in lines:
        if line.out_month != out_month or line.material_code != code:
            continue
        if line.invoice_status != SalesLedgerLine.InvoiceStatus.INVOICED:
            continue
        if line.invoice_month and line.invoice_month > line.out_month:
            continue  # 跨期开票走另一条路径
        key = (line.order_no or "", line.customer_code or "", line.customer_name or "")
        if key not in groups:
            groups[key] = {"qty": ZERO, "row_no": line.row_no}
            order.append(key)
        groups[key]["qty"] += Decimal(line.quantity)
    return [
        (key[0], key[1], key[2], groups[key]["row_no"], groups[key]["qty"]) for key in order
    ]


def _ordered_pool(pool):
    return sorted(pool, key=lambda t: (t.txn_date, t.document_no or "", t.id))


def _fill_pool(rows, qty):
    """按行剩余可核销量分配 qty，返回 ([(txn, take)], 未落实数量)。"""
    picked = []
    left = qty
    for txn in rows:
        if left <= ZERO:
            break
        capacity = abs(Decimal(txn.quantity or ZERO)) - _txn_matched_quantity(
            txn, MatchAllocation.Kind.SALES_INVOICE
        )
        if capacity <= ZERO:
            continue
        take = min(left, capacity)
        picked.append((txn, take))
        left -= take
    return picked, left


def _take_from_pool(pool, *, order_no, cust_code, cust_name, qty):
    """台账口径取量：订单号+客户 → 仅订单号 → 按物料顺序兜底。

    返回 ([(txn, take)], 未落实数量, 匹配方式说明)。
    """
    if order_no:
        exact = [
            t
            for t in pool
            if _norm(t.order_no) == _norm(order_no)
            and _same_party_strict(t.customer_code, t.customer_name, cust_code, cust_name)
        ]
        if exact:
            picked, left = _fill_pool(_ordered_pool(exact), qty)
            return picked, left, "订单号+客户"
        by_order = [t for t in pool if _norm(t.order_no) == _norm(order_no)]
        if by_order:
            picked, left = _fill_pool(_ordered_pool(by_order), qty)
            return picked, left, "仅订单号（客户不一致）"
        picked, left = _fill_pool(_ordered_pool(pool), qty)
        return picked, left, "按物料顺序兜底（台账订单号在流水里没找到）"
    picked, left = _fill_pool(_ordered_pool(pool), qty)
    return picked, left, "按物料顺序（台账未填订单号）"



def _create_invoice(period, batch, seq, *, sample=False, out_month=None,
                    invoice_month=None, order_no="", customer_code="",
                    customer_name="", material_code="", material_name="", quantity=ZERO):
    prefix = "LEDGER-SAMPLE-" if sample else "LEDGER-"
    if sample:
        invoice_no = f"{prefix}{period.month:%Y%m}-{seq:04d}"
    else:
        invoice_no = f"LEDGER-{period.month:%Y%m}-{seq:04d}"
    return SalesInvoiceLine.objects.create(
        period=period,
        batch=batch,
        invoice_no=invoice_no,
        invoice_date=invoice_month or period.month,
        invoice_period=invoice_month or period.month,
        sales_order_no=order_no or "",
        customer_code=customer_code or "",
        customer_name=customer_name or "",
        material_code=material_code,
        material_name=material_name,
        quantity=quantity,
        unit_price_excl_tax=None,
        amount_excl_tax=None,
        remaining_quantity=quantity,
    )


def _manual_alloc(period, txn, invoice, qty):
    """创建手动核销分配（样品发票→指定出库行）。"""
    MatchAllocation.objects.create(
        period=period,
        kind=MatchAllocation.Kind.SALES_INVOICE,
        inventory_transaction=txn,
        sales_invoice_line=invoice,
        matched_quantity=qty,
        is_manual=True,
        confidence=100,
    )


def apply_sales_ledger(batch: ImportBatch) -> str:
    period = batch.period
    current_month = period.month
    lines = list(batch.sales_ledger_lines.order_by("row_no"))
    if not lines:
        return "销售台账：无有效数据行。"

    notes = []
    warnings = []
    infos = []
    no_order_rows = 0

    # —— 口径汇总（出库期间×产品） ——
    sample_qty = defaultdict(Decimal)         # 样品（当月）
    invoiced_current = defaultdict(Decimal)   # 当月发货当月开票
    invoiced_prior = defaultdict(Decimal)     # 前期发货本期开票（出库期间<当月）
    uninvoiced_qty = defaultdict(Decimal)     # 未开票（当月出库挂账）

    for line in lines:
        key = (line.out_month, line.material_code)
        qty = Decimal(line.quantity)
        if line.invoice_status == SalesLedgerLine.InvoiceStatus.SAMPLE_FREE:
            if line.out_month == current_month:
                sample_qty[key] += qty
            else:
                warnings.append(
                    f"行{line.row_no} 产品 {line.material_code} 出库期间{line.out_month:%Y-%m} 的样品"
                    f"按当月处理，已跳过（样品不跨期）。"
                )
        elif line.invoice_status == SalesLedgerLine.InvoiceStatus.INVOICED:
            if line.invoice_month and line.invoice_month > line.out_month:
                invoiced_prior[key] += qty
            else:
                invoiced_current[key] += qty
        else:
            uninvoiced_qty[key] += qty

    flow_pool = _pooled_flow_txns(period)
    seq = 0

    # 1) 样品：生成 LEDGER-SAMPLE 发票并手动核销当月流水行（按日期序消耗）
    for (out_month, code), qty in sorted(sample_qty.items()):
        pool = flow_pool.get(code, [])
        available = _sum_qty(pool)
        left = qty
        # 先按“整行对应”消耗（台账一行=一张样品出库单）
        for txn in pool:
            if left <= ZERO:
                break
            row_qty = abs(Decimal(txn.quantity or ZERO))
            if row_qty <= ZERO or row_qty > left:
                continue
            seq += 1
            invoice = _create_invoice(
                period, batch, seq, sample=True,
                out_month=out_month, invoice_month=current_month,
                material_code=code, material_name=txn.material_name,
                quantity=row_qty,
            )
            _manual_alloc(period, txn, invoice, row_qty)
            left -= row_qty
        # 余量按流水行余额部分核销（台账数量与出库行不整除时的兜底，不再让整批导入失败）
        if left > ZERO:
            rest_pool = [
                t for t in flow_pool.get(code, []) if not is_sample_issue(t)
            ]
            picked, rest = _fill_pool(_ordered_pool(rest_pool), left)
            done = left - rest
            if done > ZERO:
                seq += 1
                invoice = _create_invoice(
                    period, batch, seq, sample=True,
                    out_month=out_month, invoice_month=current_month,
                    material_code=code,
                    material_name=(picked[0][0].material_name if picked else ""),
                    quantity=done,
                )
                for txn, take in picked:
                    _manual_alloc(period, txn, invoice, take)
                warnings.append(
                    f"样品 {_fmt(done)}（{code} {out_month:%Y-%m}）按出库行余额核销："
                    f"台账样品数量 {_fmt(qty)} 与出库行数量不整除，已按实际出库量计入样品费用。"
                )
                left = rest
        if left > ZERO:
            warnings.append(
                f"产品 {code} {out_month:%Y-%m}：样品数量 {_fmt(qty)} 中仍有 {_fmt(left)} "
                f"未找到对应销售出库流水（当月出库可核销量 {_fmt(available)}），"
                "未做核销，请在“销售匹配”页核对该样品出库单。"
            )
        # 刷新池：剔除已核销完的流水行，保留仍有可核销余额的行
        flow_pool[code] = [
            t
            for t in flow_pool.get(code, [])
            if abs(Decimal(t.quantity or ZERO))
            - _txn_matched_quantity(t, MatchAllocation.Kind.SALES_INVOICE)
            > ZERO
        ]
        notes.append(f"样品 {_fmt(qty)}（{code}）→ 销售费用")


    # 2) 当月已开票 → 按台账的 订单号+客户 生成结转发票并直接人工核销
    #    （台账就是匹配结果，平台不再按物料自己挑行；无订单号时按物料顺序兜底并提示）
    invoiced_total_used = ZERO
    for (out_month, code), qty in sorted(invoiced_current.items()):
        if out_month != current_month:
            invoiced_prior[(out_month, code)] += qty
            continue
        pool = flow_pool.get(code, [])
        available = _sum_qty(pool)
        if qty > available:
            warnings.append(
                f"产品 {code} {out_month:%Y-%m}：已开票数量 {_fmt(qty)} "
                f"超过当月可结转出库量 {_fmt(available)}（流水出库-样品），请核对台账。"
            )
            qty = available
        groups = _invoiced_row_groups(lines, out_month, code) or [("", "", "", None, qty)]
        for order_no, cust_code, cust_name, row_no, group_qty in groups:
            picked, left, mode = _take_from_pool(
                flow_pool.get(code, []),
                order_no=order_no,
                cust_code=cust_code,
                cust_name=cust_name,
                qty=group_qty,
            )
            done = group_qty - left
            if done <= ZERO:
                warnings.append(
                    f"产品 {code} {out_month:%Y-%m}：已开票数量 {_fmt(group_qty)} "
                    f"无法与当月出库流水对应，请核对台账与库存流水。"
                    f"无法与当月出库流水对应，已跳过该笔结转，请在“销售匹配”页人工处理。"
                )
                continue
            seq += 1
            invoice = _create_invoice(
                period, batch, seq,
                out_month=out_month, invoice_month=out_month,
                order_no=order_no, customer_code=cust_code, customer_name=cust_name,
                material_code=code, quantity=done,
            )
            for txn, take in picked:
                _manual_alloc(period, txn, invoice, take)
            invoiced_total_used += done
            tag = f" {order_no}" if order_no else ""
            notes.append(f"当月结转 {_fmt(done)}（{code}{tag}）→ 销售成本")
            if mode == "订单号+客户":
                pass
            elif mode.startswith("按物料顺序"):
                # 台账未填订单号：这是预期路径（台账已给出 期间×产品，平台只需按物料认领出库量）
                no_order_rows += 1
            else:
                warnings.append(
                    f"台账行{row_no or '-'} 产品 {code}：{mode}，已核销 {_fmt(done)}；"
                    "请在“销售匹配”页核对该订单号是否写错。"
                )
            if left > ZERO:
                warnings.append(
                    f"台账行{row_no or '-'} 产品 {code}：已开票 {_fmt(group_qty)} 中仍有 {_fmt(left)} "
                    "无法落实到出库行，暂按未开票挂发出商品，请核对台账与流水。"
                )

    # 3) 跨期已开票 → 优先核销期初发出商品明细（带订单/客户），不足部分提示人工
    for (out_month, code), qty in sorted(invoiced_prior.items()):
        if out_month == current_month:
            pool = flow_pool.get(code, [])
            available = _sum_qty(pool)
            if qty > available:
                warnings.append(
                    f"产品 {code}：已开票数量 {_fmt(qty)} 超过可结转出库量 {_fmt(available)}。"
                )
                qty = available
            seq += 1
            _create_invoice(period, batch, seq, out_month=out_month,
                            invoice_month=current_month, material_code=code,
                            quantity=qty)
            invoiced_total_used += qty
            notes.append(f"当月结转 {_fmt(qty)}（{code}）→ 销售成本")
            continue
        order_rows = defaultdict(Decimal)
        for line in lines:
            if (
                line.out_month == out_month
                and line.material_code == code
                and line.invoice_status == SalesLedgerLine.InvoiceStatus.INVOICED
                and line.invoice_month and line.invoice_month > line.out_month
            ):
                order_rows[(line.order_no or "", line.customer_name or "")] += Decimal(
                    line.quantity
                )
        for (order_no, customer_name), ord_qty in sorted(order_rows.items()):
            left = ord_qty
            opening_rows = list(
                OpeningGoodsIssued.objects.filter(
                    period__month__lte=current_month,
                    material_code=code,
                    remaining_quantity__gt=0,
                    match_mode=MatchingMode.AUTO,
                )
                .order_by("original_period", "document_date", "id")
            )
            if order_no:
                opening_rows = [
                    o for o in opening_rows if o.sales_order_no == order_no
                ]
            for opening in opening_rows:
                if left <= ZERO:
                    break
                take = min(left, Decimal(opening.remaining_quantity))
                seq += 1
                invoice = _create_invoice(
                    period, batch, seq,
                    out_month=out_month, invoice_month=current_month,
                    order_no=opening.sales_order_no or order_no,
                    customer_code=opening.customer_code,
                    customer_name=opening.customer_name or customer_name,
                    material_code=code,
                    material_name=opening.material_name or "",
                    quantity=take,
                )
                # 手动核销到该期初行，避免按订单号自动匹配顺序影响
                MatchAllocation.objects.create(
                    period=period,
                    kind=MatchAllocation.Kind.SALES_INVOICE,
                    opening_goods_issued=opening,
                    sales_invoice_line=invoice,
                    matched_quantity=take,
                    is_manual=True,
                    confidence=100,
                )
                left -= take
                invoiced_total_used += take
                notes.append(
                    f"前期结转 {_fmt(take)}（{code} {opening.sales_order_no}）→ 本期销售成本"
                )
            if left > ZERO:
                # 期初量不足：仍生成发票行（带订单），等人工核销
                seq += 1
                _create_invoice(
                    period, batch, seq,
                    out_month=out_month, invoice_month=current_month,
                    order_no=order_no, customer_name=customer_name,
                    material_code=code, quantity=left,
                )
                invoiced_total_used += left
                warnings.append(
                    f"产品 {code} {out_month:%Y-%m}：本期开票 {_fmt(left)} 未能在期初发出商品明细核销完，"
                    f"已生成 LEDGER 发票，请在“销售匹配”人工核销或核对期初发出商品明细。"
                )

    # 4) 未开票（当月出库挂账）数量对照
    for (out_month, code), qty in sorted(uninvoiced_qty.items()):
        if out_month == current_month:
            pool = flow_pool.get(code, [])
            available = _sum_qty(pool)
            if qty > available:
                warnings.append(
                    f"产品 {code} {out_month:%Y-%m}：未开票数量 {_fmt(qty)} "
                    f"超过当月流水出库剩余 {_fmt(available)}（流水-样品-已开票），请核对台账。"
                )
            notes.append(f"发出商品 {_fmt(qty)}（{code}）→ 挂账未开票")

    # 5) 当月流水总量 vs 台账覆盖对照（流水有、台账没覆盖 → 自动挂未开票，仅提示）
    ledger_cover = defaultdict(Decimal)
    for line in lines:
        if line.out_month == current_month:
            ledger_cover[line.material_code] += abs(Decimal(line.quantity))
    uncovered_codes = []
    uncovered_total = ZERO
    for code, rows in sorted(flow_pool.items()):
        flow_total = _sum_qty([t for t in rows if not is_sample_issue(t)])
        covered = ledger_cover.get(code, ZERO)
        if flow_total > covered:
            uncovered_codes.append(code)
            uncovered_total += flow_total - covered
    if uncovered_codes:
        if len(uncovered_codes) <= 12:
            for code in uncovered_codes:
                flow_total = _sum_qty(flow_pool[code])
                warnings.append(
                    f"产品 {code}：流水当月销售出库 {_fmt(flow_total)}，台账未覆盖/覆盖不足，"
                    f"差额将挂“发出商品(未开票)”；如已开票或样品请补台账行。"
                )
        else:
            warnings.append(
                f"另有 {len(uncovered_codes)} 个产品的流水销售出库未在台账覆盖，合计 {_fmt(uncovered_total)}，"
                f"将挂“发出商品(未开票)”（如 11.14 等）。如已开票/样品请补台账行。"
            )

    # 6) 执行自动核销（刷新出库行 invoice_period/match_status 及销售未开票口径）
    auto_match_all(period)

    header = (
        f"销售台账已应用：{len(lines)} 行。结转 {_fmt(invoiced_total_used)}，"
        f"发出商品 {_fmt(sum(uninvoiced_qty.values(), ZERO))}，"
        f"样品 {_fmt(sum(sample_qty.values(), ZERO))}。"
    )
    if no_order_rows:
        infos.append(
            f"本批有 {no_order_rows} 行未填订单单号，已按 产品+出库日期顺序 认领当月出库量；"
            "金额、勾稽、报表合计不受影响（台账已给出 出库期间×产品 的口径，平台不需要逐单对照）。"
        )
    parts = [header]
    parts.extend(f"  · {n}" for n in notes[:80])
    if len(notes) > 80:
        parts.append(f"  · …等共 {len(notes)} 条明细")
    parts.extend(f"  · {i}" for i in infos)
    parts.extend(f"  ⚠ {w}" for w in warnings[:40])
    if len(warnings) > 40:
        parts.append(f"  ⚠ …等共 {len(warnings)} 条提示")
    parts.append("（结转/发出商品金额在跑结算后按月末加权价计算，可在“销售未开票/对账”查看。）")
    return "\n".join(parts)


def fill_sales_ledger_costs(period):
    """结算后把台账行的 销售成本单价/销售成本/销售毛利/销售毛利率 回填。

    口径（台账只提供数量与收入，成本由平台算）：
    - 成本单价 = 该产品【本期月末加权平均价】（取自库存计价 InventoryValuation）；
    - 跨期行（出库期间早于本期）优先用期初发出商品明细上的历史单位成本，取不到再用本期加权价；
    - 销售成本 = 销售数量 × 成本单价（四舍五入到分）；
    - 销售收入 = 终端销售收入（未填则取 射阳兰邦销售收入）；
    - 毛利 = 收入 − 成本；毛利率 = 毛利 ÷ 收入（收入为空或 0 时不计算）。
    返回回填的行数。
    """
    prices = {
        v.material_code: Decimal(v.weighted_average_price or ZERO)
        for v in period.inventory_valuations.all()
    }
    updated = 0
    for line in period.sales_ledger_lines.all():
        unit = prices.get(line.material_code)
        if line.out_month and line.out_month < period.month:
            opening = (
                OpeningGoodsIssued.objects.filter(
                    material_code=line.material_code,
                    original_period=line.out_month,
                    unit_cost__isnull=False,
                )
                .order_by("id")
                .first()
            )
            if opening and opening.unit_cost is not None:
                unit = Decimal(opening.unit_cost)
        if unit is None:
            continue
        qty = Decimal(line.quantity or ZERO)
        cost = (qty * unit).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        revenue = line.terminal_amount if line.terminal_amount is not None else line.lb_amount
        line.cost_price = unit.quantize(Decimal("0.00000001"), rounding=ROUND_HALF_UP)
        line.cost_amount = cost
        if revenue is not None and Decimal(revenue) != ZERO:
            rev = Decimal(revenue)
            line.gross_profit = (rev - cost).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            line.gross_margin = ((rev - cost) / rev).quantize(
                Decimal("0.00000001"), rounding=ROUND_HALF_UP
            )
        else:
            line.gross_profit = None
            line.gross_margin = None
        line.save(
            update_fields=["cost_price", "cost_amount", "gross_profit", "gross_margin"]
        )
        updated += 1
    return updated
