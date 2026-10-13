# -*- coding: utf-8 -*-
"""生产领料（投料）凭证生成。

取数口径（用户 2026-09-17 确认）：
- **除浆粕车间外的车间**：领料按「工单投料台账」(MaterialIssue) 取数 ——
  台账行用「工单号 → 工单的生产部门」归车间；
  金额 = 实际投料数量 × 该物料月末加权平均价（InventoryValuation.weighted_average_price）。
- **浆粕车间**没有工单：按库存流水「领料出库」取数，且**备注**里必须有
  「浆粕」＋当期月份字样（如「浆粕车间7月领用」）才算浆粕车间的领料。
- 其余无工单领料（备注是别的、浆粕备注月份不是当期、或无备注）→ 归「无工单领料」。

凭证格式：
- 一张凭证、一条分录流，不分块、没有 ①②③ 小计；
- 借方科目合并：各车间合计一行「生产成本—直接材料」；
- 贷方按车间顺序（染色→纺纱→包覆→针织→机织→浆粕→无工单领料）逐行列「各车间领用的物料大类」，
  摘要「{月}月{车间}领用{物料大类}」，科目 原材料—大类／库存商品—大类；
- 附件1 逐条明细（含来源：工单投料台账 / 库存领料出库），附件2 车间 × 物料大类，附件3 贷方科目汇总。
"""
from __future__ import annotations

from collections import defaultdict
from decimal import ROUND_HALF_UP, Decimal

from settlement.models import AccountingPeriod, InventoryTransaction, WorkOrder
from settlement.services.imports import WORKSHOP_LABELS, normalize_workshop

ZERO = Decimal("0")
CENT = Decimal("0.01")

#: 借方科目（下挂车间）
ACCOUNT_DIRECT_MATERIAL = "生产成本—直接材料"
#: 贷方科目（按科目合并，不带物料大类）
ACCOUNT_RAW = "原材料"
#: 上游车间产出的自制半成品（纱线/染色纤维/熟条/坯布/面料）在账上挂「库存商品」
#: （用户 2026-09-25：原来叫「自制半成品」，改成「库存商品」）
ACCOUNT_INHOUSE = "库存商品"
#: 旧名（保留兼容，值同 ACCOUNT_INHOUSE）
ACCOUNT_SEMI = ACCOUNT_INHOUSE

#: 两种取数来源（明细里标明）
SOURCE_LEDGER = "工单投料台账"
SOURCE_FLOW = "库存领料出库"

#: 领料出库（无工单）的业务类型
PICK_TYPES = ("领料出库",)

#: 车间顺序（用户口径）：染色 → 纺纱 → 包覆 → 针织 → 机织 → 浆粕
WORKSHOP_SEQUENCE = ("RanZheng", "tefang", "BaoFu", "ZhenZhi", "JiZhi")
PULP_WORKSHOP = "浆粕车间"
#: 无法确定车间的无工单领料
UNALLOCATED = "无工单"
UNALLOCATED_LABEL = "无工单领料"

UNCLASSIFIED = "未分类"

#: 浆粕车间认领规则：库存流水备注里同时出现这两个字样才算浆粕车间领料
PULP_REMARK_KEYWORD = "浆粕"
#: 是否要求备注里出现当期月份（「7月」这种）。True＝只认当期，False＝只要含「浆粕」就算
PULP_REQUIRE_CURRENT_MONTH = True

#: 这些物料大类是上游车间产出的自制半成品（贷「库存商品」），其余按外购（贷「原材料」）。
#: 依据：本期入库来源里 纱线/染色纤维/熟条/坯布/面料 全部来自「报工」，原料/纤维 来自「采购入库」。
INHOUSE_CATEGORIES = ("纱线", "染色纤维", "熟条", "坯布", "面料")

#: 借方科目是否合并成一行（True：整张凭证只出一行「生产成本—直接材料」；
#: False：按「车间 × 物料大类」逐行展开，科目带车间后缀）
MERGE_DEBIT_WORKSHOP = True

#: 贷方是否把物料大类也去掉（False＝保留 原材料—原料 / 库存商品—纱线 这种大类后缀）
MERGE_CREDIT_CATEGORY = False

#: 摘要里的车间叫法（如需按用户习惯改，例如 "RanZheng": "染整车间"，在这里加一行即可）
SUMMARY_WORKSHOP_LABELS: dict = {}

#: 车间排序（借方行、贷方行与附件汇总表都用它）
WORKSHOP_ORDER_INDEX = {
    code: index
    for index, code in enumerate(list(WORKSHOP_SEQUENCE) + [PULP_WORKSHOP, UNALLOCATED])
}

VOUCHER_NOTE = (
    "借：生产成本—直接材料（各车间合计，一行）；"
    "贷：各车间领用的物料大类——按车间顺序分行（染色→纺纱→包覆→针织→机织→浆粕→无工单领料）。"
    "取数：除浆粕车间外按「工单投料台账」（数量×月末加权平均价，工单→生产部门归车间）；"
    "浆粕车间按库存流水「领料出库」，且备注含「浆粕 + 当期月份」；其余无工单领料单列。"
)


def money(value) -> Decimal:
    return Decimal(value or ZERO).quantize(CENT, rounding=ROUND_HALF_UP)


def _bucket():
    return {"amount": ZERO, "quantity": ZERO, "rows": 0}


def _add(bucket, amount, quantity, rows=1):
    bucket["amount"] += amount
    bucket["quantity"] += quantity
    bucket["rows"] += rows


def workshop_label(code: str) -> str:
    """车间中文显示名（摘要与科目都用它）。"""
    if not code:
        return ""
    if code == PULP_WORKSHOP:
        return PULP_WORKSHOP
    if code == UNALLOCATED:
        return UNALLOCATED_LABEL
    override = SUMMARY_WORKSHOP_LABELS.get(code)
    if override:
        return override
    return WORKSHOP_LABELS.get(code, code)


def _category_fallback(period):
    """物料分类兜底表：期初存货 → 期末存货 → 库存流水，取第一个非空分类。

    用户 2026-10-13：正式环境有一批浆粕车间领料的投料记录/流水**没有分类**，
    凭证只能归到「未分类」，和测试环境（有分类=原料）不一致。这里加兜底：
    行上没有分类时，按物料编码去这三处找，仍找不到才归「未分类」。
    """
    mapping = {}
    layers = (
        period.opening_inventory_balances.exclude(category_name=""),
        period.ending_inventory_balances.exclude(category_name=""),
        period.inventory_transactions.exclude(category_name=""),
    )
    for layer in layers:
        for code, category in layer.order_by("id").values_list("material_code", "category_name"):
            if code and category and str(category).strip():
                mapping.setdefault(code, str(category).strip())
    return mapping


def credit_account_base(category: str) -> str:
    """贷方一级科目：原材料（外购）／库存商品（上游车间产出）。"""
    return ACCOUNT_INHOUSE if category in INHOUSE_CATEGORIES else ACCOUNT_RAW


def credit_account(category: str) -> str:
    """贷方科目：原材料／库存商品，默认带物料大类后缀。"""
    base = credit_account_base(category)
    if MERGE_CREDIT_CATEGORY or not category or category == UNCLASSIFIED:
        return base
    return "%s—%s" % (base, category)


def summary_text(period, workshop, category="") -> str:
    """分录摘要：{月}月{车间}领用{物料大类}。"""
    month = "%d月" % period.month.month
    if workshop == UNALLOCATED:
        if category:
            return "%s无工单领料（%s）" % (month, category)
        return "%s无工单领料" % month
    label = workshop_label(workshop)
    if category:
        return "%s%s领用%s" % (month, label, category)
    return "%s%s领用材料" % (month, label)


def debit_summary_text(period) -> str:
    """借方摘要（借方科目已合并成一行）：{月}月各车间领用材料。"""
    return "%d月各车间领用材料" % period.month.month


def _order_departments(period):
    return {
        row.order_no: row.production_dept_code
        for row in WorkOrder.objects.filter(period=period).only(
            "order_no", "production_dept_code"
        )
    }


def _valuation_prices(period):
    return {
        row.material_code: row.weighted_average_price
        for row in period.inventory_valuations.all()
    }


def _price_of(prices, code):
    value = prices.get((code or "").strip())
    if value is None:
        return None
    return Decimal(value)


def _ledger_issue_rows(period, order_dept, prices):
    """① 工单投料台账：除浆粕车间外的车间按这个取数。

    返回 (明细行, 有数量但金额为 0 的行数, 这些行的数量, 工单找不到部门而落无工单的行数)。
    """
    rows = []
    zero_amount_rows = 0
    zero_amount_quantity = ZERO
    no_department_rows = 0
    fallback = _category_fallback(period)
    for issue in period.material_issues.all():
        order_no = (issue.work_order_no or "").strip()
        workshop = normalize_workshop(order_dept.get(order_no, ""))
        if not workshop:
            workshop = UNALLOCATED
            no_department_rows += 1
        category = (
            (issue.material_category or "").strip()
            or fallback.get(issue.material_code, "")
            or UNCLASSIFIED
        )
        quantity = Decimal(issue.actual_issue_quantity or ZERO)
        price = _price_of(prices, issue.material_code)
        if price is None:
            price = None
            amount = ZERO
        else:
            amount = money(quantity * price)
        if quantity != ZERO and amount == ZERO:
            zero_amount_rows += 1
            zero_amount_quantity += quantity
        rows.append(
            {
                "date": issue.issue_time.date() if issue.issue_time else None,
                "source": SOURCE_LEDGER,
                "document_no": "",
                "work_order_no": order_no,
                "business_type": "投料",
                "workshop": workshop,
                "workshop_label": workshop_label(workshop),
                "category": category,
                "material_code": issue.material_code,
                "material_name": issue.material_name,
                "warehouse": issue.issue_location,
                "quantity": quantity,
                "unit_price": price,
                "amount": amount,
                "basis": "工单号→生产部门"
                if workshop != UNALLOCATED else "台账工单找不到生产部门",
                "summary": summary_text(period, workshop, category),
            }
        )
    return rows, zero_amount_rows, zero_amount_quantity, no_department_rows


def _pulp_month_keyword(period) -> str:
    return "%d月" % period.month.month


def _flow_pick_rows(period):
    """② 浆粕车间＋无工单领料：库存流水「领料出库」，按备注认领浆粕车间的部分。"""
    month_kw = _pulp_month_keyword(period)
    rows = []
    pulp_rows = 0
    pulp_amount = ZERO
    mismatch_rows = 0
    mismatch_amount = ZERO
    zero_amount_rows = 0
    zero_amount_quantity = ZERO
    fallback = _category_fallback(period)
    txns = [
        txn
        for txn in period.inventory_transactions.all()
        if (txn.business_type or "").strip() in PICK_TYPES
    ]
    txns.sort(key=lambda t: (t.txn_date, t.document_no or "", t.id))
    for txn in txns:
        remark = (txn.remark or "").strip()
        category = (
            (txn.category_name or "").strip()
            or fallback.get(txn.material_code, "")
            or UNCLASSIFIED
        )
        quantity = Decimal(txn.quantity or ZERO)
        amount = money(txn.amount)
        is_pulp_remark = PULP_REMARK_KEYWORD in remark
        if is_pulp_remark and (not PULP_REQUIRE_CURRENT_MONTH or month_kw in remark):
            workshop = PULP_WORKSHOP
            basis = "流水备注含「%s + %s」" % (PULP_REMARK_KEYWORD, month_kw)
            pulp_rows += 1
            pulp_amount += amount
        elif is_pulp_remark:
            workshop = UNALLOCATED
            basis = "备注写了浆粕但月份不是当期（%s）" % remark
            mismatch_rows += 1
            mismatch_amount += amount
        else:
            workshop = UNALLOCATED
            basis = "无工单领料（备注：%s）" % remark if remark else "无工单领料（无备注）"
        if quantity != ZERO and amount == ZERO:
            zero_amount_rows += 1
            zero_amount_quantity += quantity
        rows.append(
            {
                "date": txn.txn_date,
                "source": SOURCE_FLOW,
                "document_no": txn.document_no,
                "work_order_no": "",
                "business_type": (txn.business_type or "").strip(),
                "workshop": workshop,
                "workshop_label": workshop_label(workshop),
                "category": category,
                "material_code": txn.material_code,
                "material_name": txn.material_name,
                "warehouse": txn.warehouse,
                "quantity": quantity,
                "unit_price": txn.unit_price,
                "amount": amount,
                "basis": basis,
                "remark": remark,
                "summary": summary_text(period, workshop, category),
            }
        )
    stats = {
        "pulp_rows": pulp_rows,
        "pulp_amount": money(pulp_amount),
        "mismatch_rows": mismatch_rows,
        "mismatch_amount": money(mismatch_amount),
        "zero_amount_rows": zero_amount_rows,
        "zero_amount_quantity": zero_amount_quantity,
    }
    return rows, stats


def pulp_issue_rows(period):
    """浆粕车间的领料行（按库存流水备注认领，规则与生产领料凭证一致）。

    供「工单投料」页做汇总时复用：返回 _flow_pick_rows 里归属浆粕车间的那些行。
    """
    rows, _stats = _flow_pick_rows(period)
    return [row for row in rows if row["workshop"] == PULP_WORKSHOP]


def _unmatched_flow_rows(period, ledger_rows, order_dept):
    """投料流水在工单投料台账里找不到对应记录的部分（按台账口径不计入，但要提示）。"""
    keys = defaultdict(int)
    for row in ledger_rows:
        code = (row["material_code"] or "").strip()
        keys[(row["work_order_no"], code, Decimal(row["quantity"]).quantize(Decimal("0.001")))] += 1
    unmatched = 0
    amount = ZERO
    for txn in period.inventory_transactions.all():
        if (txn.business_type or "").strip() not in {"投料", "投料回撤", "投料撤回"}:
            continue
        order_no = (txn.document_no or "").strip()
        if normalize_workshop(order_dept.get(order_no, "")) == PULP_WORKSHOP:
            continue
        code = (txn.material_code or "").strip()
        quantity = Decimal(txn.quantity or ZERO).quantize(Decimal("0.001"))
        key = (order_no, code, quantity)
        if keys.get(key):
            keys[key] -= 1
            continue
        unmatched += 1
        signed = Decimal(txn.amount or ZERO)
        amount += signed if txn.direction == InventoryTransaction.Direction.OUT else -signed
    return unmatched, money(amount)


def build_material_issue_voucher(period: AccountingPeriod, *,
                                 detail_workshop="", detail_material=""):
    """生成生产领料（投料）凭证：一张凭证、一条分录流。

    detail_workshop / detail_material 只筛选附件明细，不影响凭证金额。
    """
    order_dept = _order_departments(period)
    prices = _valuation_prices(period)

    ledger_rows, ledger_zero_rows, ledger_zero_qty, ledger_no_dept = _ledger_issue_rows(
        period, order_dept, prices
    )
    flow_rows, flow_stats = _flow_pick_rows(period)
    detail_rows = ledger_rows + flow_rows

    # ── 按 (车间, 物料大类) 汇总 ────────────────────────────────────────────
    debit = defaultdict(_bucket)
    for row in detail_rows:
        _add(
            debit[(row["workshop"], row["category"])],
            row["amount"],
            row["quantity"],
        )

    # ── 凭证分录：一张凭证、一条分录流 ──────────────────────────────────────
    #    借方：科目合并成一行「生产成本—直接材料」（各车间合计）；
    #    贷方：按车间顺序（染色→纺纱→包覆→针织→机织→浆粕→无工单领料）列各车间领用的物料大类。
    rows = []
    debit_keys = sorted(
        debit, key=lambda key: (WORKSHOP_ORDER_INDEX.get(key[0], 99), str(key[1]))
    )
    debit_total_bucket = _bucket()
    for key in debit_keys:
        item = debit[key]
        if item["amount"] == ZERO and item["quantity"] == ZERO:
            continue
        _add(debit_total_bucket, item["amount"], item["quantity"], item["rows"])
    if debit_total_bucket["amount"] != ZERO or debit_total_bucket["quantity"] != ZERO:
        if MERGE_DEBIT_WORKSHOP:
            rows.append(
                {
                    "level": "line",
                    "summary": debit_summary_text(period),
                    "account": ACCOUNT_DIRECT_MATERIAL,
                    "aux": "（%d 个车间 × %d 个物料大类，明细见附件2）" % (
                        len({key[0] for key in debit_keys}),
                        len({key[1] for key in debit_keys}),
                    ),
                    "debit": money(debit_total_bucket["amount"]),
                    "credit": ZERO,
                    "quantity": debit_total_bucket["quantity"],
                    "source_rows": debit_total_bucket["rows"],
                }
            )
        else:
            for workshop, category in debit_keys:
                item = debit[(workshop, category)]
                if item["amount"] == ZERO and item["quantity"] == ZERO:
                    continue
                rows.append(
                    {
                        "level": "line",
                        "summary": summary_text(period, workshop, category),
                        "account": "%s—%s" % (
                            ACCOUNT_DIRECT_MATERIAL, workshop_label(workshop)),
                        "aux": "",
                        "debit": money(item["amount"]),
                        "credit": ZERO,
                        "quantity": item["quantity"],
                        "source_rows": item["rows"],
                    }
                )

    for workshop, category in debit_keys:
        item = debit[(workshop, category)]
        if item["amount"] == ZERO and item["quantity"] == ZERO:
            continue
        rows.append(
            {
                "level": "line",
                "summary": summary_text(period, workshop, category),
                "account": credit_account(category),
                "aux": "",
                "debit": ZERO,
                "credit": money(item["amount"]),
                "quantity": item["quantity"],
                "source_rows": item["rows"],
            }
        )

    debit_total = sum((row["debit"] for row in rows), ZERO)
    credit_total = sum((row["credit"] for row in rows), ZERO)
    quantity_total = sum((row["quantity"] for row in rows if row["debit"]), ZERO)
    source_rows_total = 0
    for key, item in debit.items():
        source_rows_total += item["rows"]
    rows.append(
        {
            "level": "total",
            "summary": "合计",
            "account": "",
            "aux": "",
            "debit": money(debit_total),
            "credit": money(credit_total),
            "quantity": quantity_total,
            "source_rows": source_rows_total,
        }
    )

    # ── 生成前检查 ────────────────────────────────────────────────────────
    unallocated_rows = sum(
        (debit[key]["rows"] for key in debit if key[0] == UNALLOCATED), 0
    )
    unallocated_amount = sum(
        (debit[key]["amount"] for key in debit if key[0] == UNALLOCATED), ZERO
    )
    checks = []
    if ledger_zero_rows:
        checks.append(
            {
                "level": "blocking",
                "text": "工单投料台账里有 %d 行有数量但金额为 0（缺该物料的月末加权平均价，或价格本身为 0，"
                        "共 %s），本凭证会漏记，请先在「库存结转」把这些物料的价格算出来。"
                        % (ledger_zero_rows, f"{ledger_zero_qty:,.3f}"),
            }
        )
    if flow_stats["zero_amount_rows"]:
        checks.append(
            {
                "level": "blocking",
                "text": "领料出库流水里有 %d 行有数量但金额为 0（共 %s），本凭证会漏记，"
                        "请先在「库存结转」计价。" % (
                            flow_stats["zero_amount_rows"],
                            f"{flow_stats['zero_amount_quantity']:,.3f}",
                        ),
            }
        )
    if flow_stats["mismatch_rows"]:
        checks.append(
            {
                "level": "warning",
                "text": "有 %d 行领料出库的备注写了浆粕、但月份不是当期 %s（如「浆粕车间6月领用」，%s 元），"
                        "没有算作浆粕车间领料，已归入「%s」。"
                        % (flow_stats["mismatch_rows"], _pulp_month_keyword(period),
                           f"{flow_stats['mismatch_amount']:,.2f}", UNALLOCATED_LABEL),
            }
        )
    if unallocated_rows:
        checks.append(
            {
                "level": "warning",
                "text": "有 %d 行无工单领料（%s 元）没有工单、也没写「浆粕 + 当期月份」的备注"
                        "（如「郝工制衣用布」「群调令：领料倒筒」等），已归入「%s」单独列示。"
                        % (unallocated_rows, f"{unallocated_amount:,.2f}", UNALLOCATED_LABEL),
            }
        )
    if flow_stats["pulp_rows"]:
        checks.append(
            {
                "level": "info",
                "text": "浆粕车间领料 %d 行（%s 元）：按库存流水的备注认领（备注含「%s + %s」）。"
                        % (flow_stats["pulp_rows"], f"{flow_stats['pulp_amount']:,.2f}",
                           PULP_REMARK_KEYWORD, _pulp_month_keyword(period)),
            }
        )
    if ledger_no_dept:
        checks.append(
            {
                "level": "warning",
                "text": "工单投料台账里有 %d 行找不到工单的生产部门，已归入「%s」。"
                        % (ledger_no_dept, UNALLOCATED_LABEL),
            }
        )
    unmatched_rows, unmatched_amount = _unmatched_flow_rows(period, ledger_rows, order_dept)
    if unmatched_rows:
        checks.append(
            {
                "level": "info",
                "text": "另有 %d 行「投料」流水在工单投料台账里找不到对应记录（%s 元），"
                        "按「以工单投料为准」的口径没有计入本凭证。"
                        % (unmatched_rows, f"{unmatched_amount:,.2f}"),
            }
        )
    if not detail_rows:
        checks.append({"level": "warning", "text": "本期没有工单投料台账与领料出库流水。"})
    if money(debit_total - credit_total) != ZERO:
        checks.append(
            {
                "level": "blocking",
                "text": "借贷不平：借方 %s / 贷方 %s，差 %s。"
                        % (f"{debit_total:,.2f}", f"{credit_total:,.2f}",
                           f"{debit_total - credit_total:,.2f}"),
            }
        )

    # ── 附件明细（筛选只影响附件，不影响凭证金额） ────────────────────────
    workshop_key = str(detail_workshop or "").strip().lower()
    material_key = str(detail_material or "").strip().lower()

    def _keep(row):
        if workshop_key:
            haystack = "%s %s" % (row["workshop"], row["workshop_label"])
            if workshop_key not in haystack.lower():
                return False
        if material_key:
            haystack = "%s %s" % (row["material_code"] or "", row["material_name"] or "")
            if material_key not in haystack.lower():
                return False
        return True

    filtered_details = [row for row in detail_rows if _keep(row)]

    category_rows = []
    for workshop, category in debit_keys:
        item = debit[(workshop, category)]
        if item["amount"] == ZERO and item["quantity"] == ZERO:
            continue
        category_rows.append(
            {
                "workshop": workshop,
                "workshop_label": workshop_label(workshop),
                "category": category,
                "summary": summary_text(period, workshop, category),
                "quantity": item["quantity"],
                "amount": money(item["amount"]),
                "rows": item["rows"],
                "account_debit": "%s—%s" % (
                    ACCOUNT_DIRECT_MATERIAL, workshop_label(workshop))
                if not MERGE_DEBIT_WORKSHOP else ACCOUNT_DIRECT_MATERIAL,
                "account_credit": credit_account(category),
            }
        )

    workshop_rows = []
    workshop_keys = sorted(
        {key[0] for key in debit}, key=lambda code: WORKSHOP_ORDER_INDEX.get(code, 99)
    )
    for workshop in workshop_keys:
        total = _bucket()
        categories = 0
        for key in debit:
            if key[0] != workshop:
                continue
            item = debit[key]
            if item["amount"] == ZERO and item["quantity"] == ZERO:
                continue
            categories += 1
            _add(total, item["amount"], item["quantity"], item["rows"])
        if not categories:
            continue
        workshop_rows.append(
            {
                "workshop": workshop,
                "workshop_label": workshop_label(workshop),
                "amount": money(total["amount"]),
                "quantity": total["quantity"],
                "rows": total["rows"],
                "categories": categories,
            }
        )

    credit_rows = []
    credit_by_account = defaultdict(_bucket)
    credit_category_by_account = {}
    for workshop, category in debit_keys:
        item = debit[(workshop, category)]
        account = credit_account(category)
        credit_category_by_account[account] = category
        _add(credit_by_account[account], item["amount"], item["quantity"], item["rows"])
    for account, item in sorted(
        credit_by_account.items(),
        key=lambda pair: (pair[0].startswith(ACCOUNT_INHOUSE), pair[0]),
    ):
        if item["amount"] == ZERO and item["quantity"] == ZERO:
            continue
        credit_rows.append(
            {
                "account": account,
                "category": credit_category_by_account.get(account, ""),
                "summary": "%s各车间领用%s" % (
                    _pulp_month_keyword(period), credit_category_by_account.get(account, "")),
                "quantity": item["quantity"],
                "amount": money(item["amount"]),
                "rows": item["rows"],
            }
        )

    return {
        "month": period.month.strftime("%Y-%m"),
        "rows": rows,
        "totals": {
            "debit": money(debit_total),
            "credit": money(credit_total),
            "difference": money(debit_total - credit_total),
            "balanced": money(debit_total - credit_total) == ZERO,
            "quantity": quantity_total,
            "line_count": len([row for row in rows if row["level"] == "line"]),
        },
        "summary": {
            "ledger_rows": len(ledger_rows),
            "ledger_amount": money(sum((row["amount"] for row in ledger_rows), ZERO)),
            "pulp_rows": flow_stats["pulp_rows"],
            "pulp_amount": flow_stats["pulp_amount"],
            "flow_rows": len(flow_rows),
            "unallocated_rows": unallocated_rows,
            "unallocated_amount": money(unallocated_amount),
            "mismatch_rows": flow_stats["mismatch_rows"],
            "mismatch_amount": flow_stats["mismatch_amount"],
            "voucher_amount": money(debit_total),
            "row_count": len(detail_rows),
            "zero_amount_rows": ledger_zero_rows + flow_stats["zero_amount_rows"],
            "workshop_count": len([
                row for row in workshop_rows if row["workshop"] != UNALLOCATED
            ]),
        },
        "checks": checks,
        "headers": ["摘要", "科目", "数量", "借方", "贷方"],
        "voucher_note": VOUCHER_NOTE,
        "detail_rows": filtered_details,
        "category_rows": category_rows,
        "workshop_rows": workshop_rows,
        "credit_rows": credit_rows,
        "accounts": {
            "direct_material": ACCOUNT_DIRECT_MATERIAL,
            "raw": ACCOUNT_RAW,
            "inhouse": ACCOUNT_INHOUSE,
            "semi": ACCOUNT_INHOUSE,  # 兼容旧模板变量名
        },
        "months": [
            {"value": item.month.strftime("%Y-%m"), "label": str(item)}
            for item in AccountingPeriod.objects.order_by("month")
        ],
        "filter": {
            "workshop": detail_workshop or "",
            "material": detail_material or "",
            "active": bool(workshop_key or material_key),
        },
    }
