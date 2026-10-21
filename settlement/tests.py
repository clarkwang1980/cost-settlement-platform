import re

from datetime import date, datetime
from decimal import Decimal
from io import BytesIO, StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace

from django.conf import settings
from django.contrib.auth import get_user_model
from django.core.files.uploadedfile import SimpleUploadedFile
from django.core.management import call_command
from django.test import TestCase, override_settings
from django.urls import reverse
from django.utils import timezone
from openpyxl import load_workbook

from settlement.models import (
    AccountingPeriod,
    BomLine,
    EndingInventoryBalance,
    HistoricalPrice,
    ImportBatch,
    IssueProgress,
    InventoryTransaction,
    InventoryValuation,
    MatchAllocation,
    MatchingMode,
    Material,
    MaterialIssue,
    OpeningGoodsIssued,
    OpeningInventoryBalance,
    OpeningPurchaseAccrual,
    OpeningWip,
    PurchaseContractLine,
    PurchaseInvoiceLine,
    PulpMaterial,
    ReconciliationResult,
    ProfitStatementExpense,
    SalesInvoiceLine,
    SalesLedgerLine,
    SettlementRun,
    StandardCostRate,
    WorkOrder,
    WorkOrderCostResult,
    WorkOrderDetail,
    WorkReport,
    WorkshopSettlement,
    WorkshopCostPool,
)
from settlement.services.costing import (
    can_lock,
    ensure_workshop_settlements,
    run_settlement,
    settle_workshop,
    update_workshop_sequences,
)
from settlement.services.imports import create_import_batch, process_batch
from settlement.services.matching import (
    auto_match_all,
    manual_match,
    set_source_match_mode,
)
from settlement import views as views
from settlement.services.receipt_voucher import build_finished_receipt_voucher
from settlement.services.material_issue_voucher import (
    PULP_WORKSHOP,
    UNALLOCATED,
    build_material_issue_voucher,
)
from settlement.services.purchase_voucher import build_purchase_voucher
from settlement.services.voucher_center import build_voucher_center
from settlement.services.purchase_accruals import (
    build_purchase_accrual_rows,
    build_purchase_difference_rows,
    provisional_basis,
)
from settlement.services.workorder_costing import (
    CASE_CHOICES,
    CASE_NO_PRICE,
    CASE_NO_REPORT,
    CASE_NORMAL,
    CASE_OTHER,
    CASE_REPORT_ONLY,
    CASE_SUGGESTIONS,
    build_work_order_cost_summary,
)
from settlement.services.pulp_costing import collect_pulp_costs
from settlement.services.sales_voucher import (
    build_sales_cost_statistics,
    build_sales_voucher_entries,
    build_sales_voucher_summary,
)
from settlement.services.sales_uninvoiced import (
    build_goods_issued_rollforward,
    build_current_uninvoiced_rows,
    build_sales_uninvoiced_rows,
    summarize_current_sales_cost,
)
from settlement.services.profit_statement import (
    ALLOCATION_FIELDS,
    HEADERS,
    build_category_summary,
    category_label,
    detail_month_options,
    filter_detail_rows,
    row_totals,
    build_profit_statement,
    build_summary,
    summary_category_options,
    summary_period_options,
)


class SettlementFixtureMixin:
    def create_batch(self, period, kind, suffix):
        return ImportBatch.objects.create(
            period=period,
            kind=kind,
            file=f"tests/{suffix}.csv",
            original_name=f"{suffix}.csv",
            sha256=(suffix * 64)[:64],
            status=ImportBatch.Status.SUCCESS,
            row_count=1,
            uploaded_by=self.user,
        )

    def create_txn(
        self,
        period,
        batch,
        *,
        txn_date,
        document_no,
        business_type,
        direction,
        material,
        quantity,
        order_no="",
        workshop="",
        supplier_code="",
        supplier_name="",
        customer_code="",
        customer_name="",
        contract_no="",
    ):
        return InventoryTransaction.objects.create(
            period=period,
            batch=batch,
            txn_date=txn_date,
            document_no=document_no,
            line_no=str(
                InventoryTransaction.objects.filter(period=period).count() + 1
            ),
            business_type=business_type,
            direction=direction,
            warehouse="测试仓库",
            material=material,
            material_code=material.code,
            material_name=material.name,
            quantity=quantity,
            uom=material.uom,
            order_no=order_no,
            workshop=workshop,
            supplier_code=supplier_code,
            supplier_name=supplier_name,
            customer_code=customer_code,
            customer_name=customer_name,
            contract_no=contract_no,
        )


class SettlementEngineTests(SettlementFixtureMixin, TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="tester", password="test-password"
        )

    def test_complete_purchase_order_inventory_and_sales_close(self):
        period = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        inventory_batch = self.create_batch(
            period, ImportBatch.Kind.INVENTORY, "inventory-a"
        )
        contract_batch = self.create_batch(
            period, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-a"
        )
        purchase_invoice_batch = self.create_batch(
            period, ImportBatch.Kind.PURCHASE_INVOICE, "purchase-invoice-a"
        )
        sales_invoice_batch = self.create_batch(
            period, ImportBatch.Kind.SALES_INVOICE, "sales-invoice-a"
        )
        order_batch = self.create_batch(
            period, ImportBatch.Kind.WORK_ORDER, "order-a"
        )
        cost_batch = self.create_batch(
            period, ImportBatch.Kind.WORKSHOP_COST, "cost-a"
        )
        raw = Material.objects.create(code="RM-A", name="原料A", uom="kg")
        product = Material.objects.create(code="FG-A", name="产品A", uom="kg")

        PurchaseContractLine.objects.create(
            period=period,
            batch=contract_batch,
            contract_no="PC-A",
            contract_date=date(2026, 6, 1),
            supplier_code="SUP-A",
            supplier_name="供应商A",
            material_code=raw.code,
            material_name=raw.name,
            quantity=Decimal("100"),
            remaining_quantity=Decimal("100"),
            unit_price_excl_tax=Decimal("10"),
        )
        PurchaseInvoiceLine.objects.create(
            period=period,
            batch=purchase_invoice_batch,
            invoice_no="PI-A",
            invoice_date=date(2026, 6, 25),
            invoice_period=date(2026, 6, 1),
            contract_no="PC-A",
            supplier_code="SUP-A",
            supplier_name="供应商A",
            material_code=raw.code,
            material_name=raw.name,
            quantity=Decimal("60"),
            remaining_quantity=Decimal("60"),
            unit_price_excl_tax=Decimal("12"),
            amount_excl_tax=Decimal("720"),
        )
        SalesInvoiceLine.objects.create(
            period=period,
            batch=sales_invoice_batch,
            invoice_no="SI-A",
            invoice_date=date(2026, 6, 28),
            invoice_period=date(2026, 6, 1),
            sales_order_no="SO-A",
            customer_code="CUS-A",
            customer_name="客户A",
            material_code=product.code,
            material_name=product.name,
            quantity=Decimal("4"),
            remaining_quantity=Decimal("4"),
        )
        order = WorkOrder.objects.create(
            period=period,
            batch=order_batch,
            order_no="WO-A",
            production_dept_code="染整",
            plan_quantity=Decimal("10"),
        )
        MaterialIssue.objects.create(
            period=period,
            batch=inventory_batch,
            work_order_no=order.order_no,
            material_code=raw.code,
            material_name=raw.name,
            issue_quantity=Decimal("50"),
            actual_issue_quantity=Decimal("50"),
            issue_uom="kg",
        )
        WorkReport.objects.create(
            period=period,
            batch=inventory_batch,
            work_order_no=order.order_no,
            material_code=product.code,
            material_name=product.name,
            quantity=Decimal("10"),
        )
        WorkshopCostPool.objects.create(
            period=period,
            batch=cost_batch,
            workshop="染整",
            water_power=Decimal("100"),
            wages=Decimal("50"),
        )
        purchase = self.create_txn(
            period,
            inventory_batch,
            txn_date=date(2026, 6, 5),
            document_no="GR-A",
            business_type="采购入库",
            direction=InventoryTransaction.Direction.IN,
            material=raw,
            quantity=Decimal("100"),
            supplier_code="SUP-A",
            supplier_name="供应商A",
            contract_no="PC-A",
        )
        self.create_txn(
            period,
            inventory_batch,
            txn_date=date(2026, 6, 10),
            document_no="MI-A",
            business_type="投料",
            direction=InventoryTransaction.Direction.OUT,
            material=raw,
            quantity=Decimal("50"),
            order_no=order.order_no,
            workshop="染整",
        )
        production_receipt = self.create_txn(
            period,
            inventory_batch,
            txn_date=date(2026, 6, 20),
            document_no=order.order_no,
            business_type="报工",
            direction=InventoryTransaction.Direction.IN,
            material=product,
            quantity=Decimal("10"),
            order_no=order.order_no,
            workshop="染整",
        )
        sales_issue = self.create_txn(
            period,
            inventory_batch,
            txn_date=date(2026, 6, 24),
            document_no="SO-A",
            business_type="销售出库",
            direction=InventoryTransaction.Direction.OUT,
            material=product,
            quantity=Decimal("4"),
            order_no="SO-A",
            customer_code="CUS-A",
            customer_name="客户A",
        )

        run = run_settlement(period)
        self.assertTrue(run.converged)
        self.assertLessEqual(run.iterations, 5)

        purchase.refresh_from_db()
        self.assertEqual(purchase.amount, Decimal("1000.00"))
        self.assertEqual(purchase.unit_price, Decimal("10.00000000"))
        self.assertEqual(purchase.price_source, "合同暂估（部分到票）")
        self.assertEqual(purchase.invoice_period, date(2026, 6, 1))

        result = order.cost_result
        self.assertEqual(result.direct_material, Decimal("500.00"))
        self.assertEqual(result.water_power, Decimal("100.00"))
        self.assertEqual(result.wages, Decimal("50.00"))
        self.assertEqual(result.finished_cost, Decimal("650.00"))
        self.assertEqual(result.wip_cost, Decimal("0.00"))

        production_receipt.refresh_from_db()
        sales_issue.refresh_from_db()
        self.assertEqual(production_receipt.amount, Decimal("650.00"))
        self.assertEqual(production_receipt.unit_price, Decimal("65.00000000"))
        self.assertEqual(sales_issue.amount, Decimal("260.00"))
        # 方案A（2026-09-21）：销售出库不再参与匹配 → 不再有「到票期间」
        self.assertIsNone(sales_issue.invoice_period)

        self.assertFalse(
            period.reconciliations.filter(
                status=ReconciliationResult.Status.FAIL
            ).exists()
        )
        lockable, reason = can_lock(period)
        self.assertTrue(lockable, reason)

        self.client.force_login(self.user)
        export_response = self.client.get(
            reverse("export_period", kwargs={"period_id": period.id})
        )
        self.assertEqual(export_response.status_code, 200)
        workbook = load_workbook(BytesIO(export_response.content), read_only=True)
        self.assertEqual(
            workbook.sheetnames,
            [
                "库存计价",
                "跨期结转",
                "期初采购暂估",
                "期初发出商品",
                "期末存货",
                "车间结转",
                "库存流水",
                "采购匹配",
                "月末采购暂估",
                "采购暂估差异",
                "销售匹配",
                "月末已出库未开票",
                "发出商品滚动",
                "工单成本",
                "勾稽结果",
                "异常清单",
            ],
        )
        workbook.close()

        lock_response = self.client.post(
            reverse("lock_period", kwargs={"period_id": period.id})
        )
        self.assertEqual(lock_response.status_code, 302)
        period.refresh_from_db()
        self.assertEqual(period.status, AccountingPeriod.Status.LOCKED)

    def test_goods_issued_follows_sales_ledger_after_plan_a(self):
        """方案A：发出商品/未开票只认销售台账的「结转类型」，不再由销售发票核销库存流水。"""
        july = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        ledger_batch = self.create_batch(
            july, ImportBatch.Kind.SALES_LEDGER, "ledger-plan-a"
        )
        goods_batch = self.create_batch(
            july, ImportBatch.Kind.OPENING_GOODS_ISSUED, "opening-plan-a"
        )
        inventory_batch = self.create_batch(
            july, ImportBatch.Kind.INVENTORY, "inv-plan-a"
        )
        invoice_batch = self.create_batch(
            july, ImportBatch.Kind.SALES_INVOICE, "invoice-plan-a"
        )
        product = Material.objects.create(code="FG-PLAN-A", name="方案A产品", uom="件")
        # 期初发出商品（上传）
        OpeningGoodsIssued.objects.create(
            period=july, batch=goods_batch, original_period=date(2026, 5, 1),
            document_date=date(2026, 5, 20), document_no="OG-PLAN-A", line_no="1",
            sales_order_no="SO-OG", customer_name="期初客户", material_code="FG-OG",
            material_name="期初产品", uom="件", original_quantity=Decimal("12"),
            prior_invoiced_quantity=Decimal("2"), opening_quantity=Decimal("10"),
            unit_cost=Decimal("100"), opening_amount=Decimal("1000"),
            remaining_quantity=Decimal("10"),
        )
        # 台账四类结转
        def ledger_line(row_no, carry, qty, cost):
            return SalesLedgerLine.objects.create(
                period=july, batch=ledger_batch, row_no=row_no,
                out_month=july.month, invoice_status=(
                    SalesLedgerLine.InvoiceStatus.INVOICED
                    if "开票结转" in carry
                    else SalesLedgerLine.InvoiceStatus.UNINVOICED
                ),
                invoice_month=july.month if "开票结转" in carry else None,
                order_no="SO-%d" % row_no, customer_name="客户%d" % row_no,
                sales_category="成品", material=product,
                material_code=product.code, material_name=product.name,
                uom="件", quantity=Decimal(qty), cost_amount=Decimal(cost),
                carry_type=carry,
            )

        ledger_line(1, "当月发货未开票挂发出商品", "6", "600")
        ledger_line(2, "前期发货本期开票结转成本", "4", "400")
        ledger_line(3, "当月发货当月开票结转成本", "2", "200")
        ledger_line(4, "免费样品赠送计入销售费用", "1", "50")
        # 方案A（用户 2026-09-27）：台账里同样带出来的「以前月份出库」未开票行，
        # 属于期初那批货 → 不再计入「本期新增发出商品」，避免重复计量。
        prior_uninvoiced = SalesLedgerLine.objects.create(
            period=july, batch=ledger_batch, row_no=5,
            out_month=date(2026, 5, 1),
            invoice_status=SalesLedgerLine.InvoiceStatus.UNINVOICED,
            order_no="SO-5", customer_name="期初客户", sales_category="成品",
            material=product, material_code=product.code, material_name=product.name,
            uom="件", quantity=Decimal("10"), cost_amount=Decimal("1000"),
            carry_type="当月发货未开票挂发出商品",
        )

        # 库存流水里放一条销售出库 + 一张销售发票：方案A 下不应产生任何核销
        sales_issue = self.create_txn(
            july, inventory_batch, txn_date=date(2026, 7, 10),
            document_no="SO-PLAN-A", business_type="销售出库",
            direction=InventoryTransaction.Direction.OUT, material=product,
            quantity=Decimal("6"), order_no="SO-1", customer_name="客户1",
        )
        SalesInvoiceLine.objects.create(
            period=july, batch=invoice_batch, invoice_no="SI-PLAN-A",
            invoice_date=date(2026, 7, 20), invoice_period=date(2026, 7, 1),
            sales_order_no="SO-1", customer_name="客户1",
            material_code=product.code, material_name=product.name,
            quantity=Decimal("6"), remaining_quantity=Decimal("6"),
        )
        auto_match_all(july)
        sales_issue.refresh_from_db()
        self.assertFalse(
            MatchAllocation.objects.filter(
                kind=MatchAllocation.Kind.SALES_INVOICE
            ).exists()
        )
        self.assertIsNone(sales_issue.invoice_period)
        self.assertEqual(
            SalesInvoiceLine.objects.get(invoice_no="SI-PLAN-A").remaining_quantity,
            Decimal("6"),
        )

        # 明细表：列出月末仍未开票的全部行（含以前月份出库那条，它就在 5 月出库）
        rows = build_sales_uninvoiced_rows(july)
        self.assertEqual(len(rows), 2)
        by_order = {row.order_no: row for row in rows}
        self.assertEqual(by_order["SO-1"].cost_amount, Decimal("600"))
        self.assertEqual(by_order["SO-1"].quantity, Decimal("6"))
        self.assertEqual(by_order["SO-5"].cost_amount, Decimal("1000"))
        self.assertTrue(by_order["SO-5"].prior_period)
        self.assertFalse(by_order["SO-1"].prior_period)

        # 本期新增只算本期出库的（SO-1 那 600）；5 月那条＝期初那批，不重复计入
        current_rows = build_current_uninvoiced_rows(july)
        self.assertEqual(len(current_rows), 1)
        self.assertEqual(current_rows[0].order_no, "SO-1")

        # 发出商品滚动：期初 1000 ＋ 本期新增 600 － 前期开票结转 400 ＝ 1200
        rollforward = build_goods_issued_rollforward(july)
        self.assertEqual(rollforward["opening_amount"], Decimal("1000.00"))
        self.assertEqual(rollforward["current_issue_amount"], Decimal("600.00"))
        self.assertEqual(rollforward["invoice_cleared_amount"], Decimal("400.00"))
        self.assertEqual(rollforward["ending_amount"], Decimal("1200.00"))
        self.assertEqual(rollforward["difference"], Decimal("0.00"))

        # 勾稽口径：台账成本合计 ＝ 已开票结转 ＋ 未开票 ＋ 样品
        costs = summarize_current_sales_cost(july)
        self.assertEqual(costs["sales_cost"], Decimal("1250.00"))
        self.assertEqual(costs["invoiced_cost"], Decimal("600.00"))
        self.assertEqual(costs["goods_issued"], Decimal("600.00"))
        self.assertEqual(costs["sample_cost"], Decimal("50.00"))

        # 页面：台账口径的列 + 数字
        self.client.force_login(self.user)
        page = self.client.get(
            reverse("sales_uninvoiced", kwargs={"period_id": july.id})
        )
        html = page.content.decode("utf-8")
        for label in ("发出商品跨月滚动勾稽", "发出商品及销售成本统计表",
                      "截至月末仍未开票的销售出库明细"):
            self.assertIn(label, html)
        for header in ("出库期间", "销售订单号", "客户", "销售成本", "未开票天数"):
            self.assertIn(header, html)
        self.assertIn("1,200.00", html)
        self.assertIn("销售台账", html)

        # 导出：表头＝台账列
        resp = self.client.get(
            reverse("export_listing", kwargs={"period_id": july.id, "kind": "sales_uninvoiced"})
        )
        wb = load_workbook(BytesIO(resp.content), read_only=True)
        try:
            ws = wb["月末未开票"]
            heads = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
            rows = list(ws.iter_rows(min_row=2, values_only=True))
        finally:
            wb.close()
        self.assertEqual(heads[0], "出库期间")
        self.assertIn("销售成本", heads)
        self.assertEqual(len(rows), 2)
        self.assertEqual(
            {row[heads.index("销售订单号")] for row in rows}, {"SO-1", "SO-5"}
        )
    def test_sales_matching_is_gone_after_plan_a(self):
        """方案A：销售出库不再参与匹配 —— 手工核销销售发票会被明确拒绝，也不产生分摊。"""
        period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        inventory_batch = self.create_batch(period, ImportBatch.Kind.INVENTORY, "no-sales-match")
        invoice_batch = self.create_batch(period, ImportBatch.Kind.SALES_INVOICE, "no-sales-inv")
        product = Material.objects.create(code="FG-NO-MATCH", name="不核销产品", uom="件")
        txn = self.create_txn(
            period, inventory_batch, txn_date=date(2026, 7, 5),
            document_no="SO-NO-MATCH", business_type="销售出库",
            direction=InventoryTransaction.Direction.OUT, material=product,
            quantity=Decimal("5"), order_no="SO-NM", customer_name="客户NM",
        )
        invoice = SalesInvoiceLine.objects.create(
            period=period, batch=invoice_batch, invoice_no="SI-NO-MATCH",
            invoice_date=date(2026, 7, 6), invoice_period=date(2026, 7, 1),
            sales_order_no="SO-NM", customer_name="客户NM",
            material_code=product.code, material_name=product.name,
            quantity=Decimal("5"), remaining_quantity=Decimal("5"),
        )
        with self.assertRaisesMessage(ValueError, "销售出库不再参与匹配"):
            manual_match(
                period,
                MatchAllocation.Kind.SALES_INVOICE,
                txn.id,
                invoice.id,
                Decimal("5"),
            )
        self.assertFalse(
            MatchAllocation.objects.filter(
                kind=MatchAllocation.Kind.SALES_INVOICE
            ).exists()
        )
        # 库存流水匹配页：销售出库显示为「无需匹配」
        self.client.force_login(self.user)
        page = self.client.get(
            reverse("inventory_matches", kwargs={"period_id": period.id}),
            {"document_no": "SO-NO-MATCH"},
        )
        self.assertContains(page, "无需匹配")
    def test_cross_period_invoice_clears_accrual_without_repricing_inventory(self):
        june = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        july = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        inventory_batch = self.create_batch(
            june, ImportBatch.Kind.INVENTORY, "inventory-accrual"
        )
        contract_batch = self.create_batch(
            june, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-accrual"
        )
        june_invoice_batch = self.create_batch(
            june, ImportBatch.Kind.PURCHASE_INVOICE, "invoice-accrual-june"
        )
        july_invoice_batch = self.create_batch(
            july, ImportBatch.Kind.PURCHASE_INVOICE, "invoice-accrual-july"
        )
        material = Material.objects.create(
            code="RM-ACCRUAL", name="暂估原料", uom="kg"
        )
        PurchaseContractLine.objects.create(
            period=june,
            batch=contract_batch,
            contract_no="PC-ACCRUAL",
            contract_date=date(2026, 6, 1),
            supplier_code="SUP-ACCRUAL",
            supplier_name="暂估供应商",
            material_code=material.code,
            material_name=material.name,
            quantity=Decimal("100"),
            remaining_quantity=Decimal("100"),
            unit_price_excl_tax=Decimal("10"),
        )
        PurchaseInvoiceLine.objects.create(
            period=june,
            batch=june_invoice_batch,
            invoice_no="PI-ACCRUAL-06",
            invoice_date=date(2026, 6, 25),
            invoice_period=date(2026, 6, 1),
            contract_no="PC-ACCRUAL",
            supplier_code="SUP-ACCRUAL",
            supplier_name="暂估供应商",
            material_code=material.code,
            material_name=material.name,
            quantity=Decimal("60"),
            remaining_quantity=Decimal("60"),
            unit_price_excl_tax=Decimal("12"),
            amount_excl_tax=Decimal("720"),
        )
        receipt = self.create_txn(
            june,
            inventory_batch,
            txn_date=date(2026, 6, 5),
            document_no="GR-ACCRUAL",
            business_type="采购入库",
            direction=InventoryTransaction.Direction.IN,
            material=material,
            quantity=Decimal("100"),
            supplier_code="SUP-ACCRUAL",
            supplier_name="暂估供应商",
            contract_no="PC-ACCRUAL",
        )

        auto_match_all(june)
        receipt.refresh_from_db()
        self.assertEqual(receipt.unit_price, Decimal("10.00000000"))
        self.assertEqual(receipt.amount, Decimal("1000.00"))
        june_difference = build_purchase_difference_rows(june)
        self.assertEqual(len(june_difference), 1)
        self.assertEqual(
            june_difference[0].difference_amount, Decimal("120.00")
        )
        june_accrual = build_purchase_accrual_rows(june)
        self.assertEqual(len(june_accrual), 1)
        self.assertEqual(
            june_accrual[0].outstanding_quantity, Decimal("40")
        )
        self.assertEqual(
            june_accrual[0].provisional_amount, Decimal("400.00")
        )

        june.status = AccountingPeriod.Status.LOCKED
        june.save(update_fields=["status"])
        PurchaseInvoiceLine.objects.create(
            period=july,
            batch=july_invoice_batch,
            invoice_no="PI-ACCRUAL-07",
            invoice_date=date(2026, 7, 20),
            invoice_period=date(2026, 7, 1),
            contract_no="PC-ACCRUAL",
            supplier_code="SUP-ACCRUAL",
            supplier_name="暂估供应商",
            material_code=material.code,
            material_name=material.name,
            quantity=Decimal("40"),
            remaining_quantity=Decimal("40"),
            unit_price_excl_tax=Decimal("11"),
            amount_excl_tax=Decimal("440"),
        )

        auto_match_all(july)
        receipt.refresh_from_db()
        self.assertEqual(receipt.unit_price, Decimal("10.00000000"))
        self.assertEqual(receipt.amount, Decimal("1000.00"))
        self.assertEqual(receipt.invoice_period, date(2026, 7, 1))
        july_difference = build_purchase_difference_rows(july)
        self.assertEqual(len(july_difference), 1)
        self.assertTrue(july_difference[0].cross_period)
        self.assertEqual(
            july_difference[0].difference_amount, Decimal("40.00")
        )
        self.assertEqual(build_purchase_accrual_rows(july), [])

        self.client.force_login(self.user)
        accrual_response = self.client.get(
            reverse("purchase_accruals", kwargs={"period_id": june.id})
        )
        self.assertContains(accrual_response, "400.00")
        # 「本月采购暂估差异」已并入月末暂估页（2026-09-21）：旧地址只做跳转
        difference_response = self.client.get(
            reverse("purchase_differences", kwargs={"period_id": july.id})
        )
        self.assertEqual(difference_response.status_code, 302)
        merged_response = self.client.get(
            reverse("purchase_accruals", kwargs={"period_id": july.id})
        )
        self.assertContains(merged_response, "跨期核销")
        self.assertContains(merged_response, "40.00")

    def test_first_period_opening_ledgers_are_cleared_by_current_invoices(self):
        period = AccountingPeriod.objects.create(month=date(2026, 1, 1))
        purchase_opening_batch = self.create_batch(
            period,
            ImportBatch.Kind.OPENING_PURCHASE_ACCRUAL,
            "opening-purchase-ledger",
        )
        goods_opening_batch = self.create_batch(
            period,
            ImportBatch.Kind.OPENING_GOODS_ISSUED,
            "opening-goods-ledger",
        )
        purchase_invoice_batch = self.create_batch(
            period,
            ImportBatch.Kind.PURCHASE_INVOICE,
            "opening-purchase-invoice",
        )
        sales_invoice_batch = self.create_batch(
            period,
            ImportBatch.Kind.SALES_INVOICE,
            "opening-sales-invoice",
        )
        purchase_opening = OpeningPurchaseAccrual.objects.create(
            period=period,
            batch=purchase_opening_batch,
            original_period=date(2025, 12, 1),
            document_date=date(2025, 12, 20),
            document_no="OPEN-GR-001",
            line_no="1",
            supplier_code="SUP-OPEN",
            supplier_name="期初供应商",
            contract_no="PC-OPEN",
            material_code="RM-OPEN",
            material_name="期初原料",
            uom="kg",
            original_quantity=Decimal("120"),
            prior_invoiced_quantity=Decimal("20"),
            opening_quantity=Decimal("100"),
            provisional_unit_price=Decimal("10"),
            opening_amount=Decimal("1000"),
            remaining_quantity=Decimal("100"),
        )
        goods_opening = OpeningGoodsIssued.objects.create(
            period=period,
            batch=goods_opening_batch,
            original_period=date(2025, 12, 1),
            document_date=date(2025, 12, 22),
            document_no="OPEN-SO-001",
            line_no="1",
            sales_order_no="SO-OPEN",
            customer_code="CUS-OPEN",
            customer_name="期初客户",
            material_code="FG-OPEN",
            material_name="期初产品",
            uom="件",
            original_quantity=Decimal("12"),
            prior_invoiced_quantity=Decimal("2"),
            opening_quantity=Decimal("10"),
            unit_cost=Decimal("100"),
            opening_amount=Decimal("1000"),
            remaining_quantity=Decimal("10"),
        )
        purchase_invoice = PurchaseInvoiceLine.objects.create(
            period=period,
            batch=purchase_invoice_batch,
            invoice_no="PI-OPEN-001",
            invoice_date=date(2026, 1, 18),
            invoice_period=date(2026, 1, 1),
            contract_no="PC-OPEN",
            supplier_code="SUP-OPEN",
            supplier_name="期初供应商",
            material_code="RM-OPEN",
            material_name="期初原料",
            quantity=Decimal("40"),
            remaining_quantity=Decimal("40"),
            unit_price_excl_tax=Decimal("12"),
            amount_excl_tax=Decimal("480"),
        )
        sales_invoice = SalesInvoiceLine.objects.create(
            period=period,
            batch=sales_invoice_batch,
            invoice_no="SI-OPEN-001",
            invoice_date=date(2026, 1, 20),
            invoice_period=date(2026, 1, 1),
            sales_order_no="SO-OPEN",
            customer_code="CUS-OPEN",
            customer_name="期初客户",
            material_code="FG-OPEN",
            material_name="期初产品",
            quantity=Decimal("4"),
            remaining_quantity=Decimal("4"),
        )

        auto_match_all(period)
        purchase_opening.refresh_from_db()
        goods_opening.refresh_from_db()
        purchase_invoice.refresh_from_db()
        sales_invoice.refresh_from_db()
        self.assertEqual(period.inventory_transactions.count(), 0)
        self.assertEqual(purchase_invoice.remaining_quantity, Decimal("0"))
        # 方案A：销售发票不再被核销 → 剩余数量保持发票数量
        self.assertEqual(
            sales_invoice.remaining_quantity, sales_invoice.quantity
        )
        self.assertEqual(purchase_opening.remaining_quantity, Decimal("60"))
        # 期初发出商品不再被销售核销消耗
        self.assertEqual(
            goods_opening.remaining_quantity, goods_opening.opening_quantity
        )

        accrual_rows = build_purchase_accrual_rows(period)
        self.assertEqual(len(accrual_rows), 1)
        self.assertIsNone(accrual_rows[0].transaction)
        self.assertEqual(
            accrual_rows[0].provisional_amount,
            Decimal("600.00"),
        )
        difference_rows = build_purchase_difference_rows(period)
        self.assertEqual(len(difference_rows), 1)
        self.assertEqual(
            difference_rows[0].difference_amount,
            Decimal("80.00"),
        )
        # 未开票/发出商品改按销售台账口径：本用例没有台账行 → 没有未开票明细，
        # 期末发出商品＝期初发出商品（上传）1000
        sales_rows = build_sales_uninvoiced_rows(period)
        self.assertEqual(sales_rows, [])
        rollforward = build_goods_issued_rollforward(period)
        self.assertEqual(rollforward["opening_amount"], Decimal("1000.00"))
        self.assertEqual(rollforward["current_issue_amount"], Decimal("0.00"))
        self.assertEqual(rollforward["invoice_cleared_amount"], Decimal("0.00"))
        self.assertEqual(rollforward["ending_amount"], Decimal("1000.00"))
        self.assertEqual(rollforward["difference"], Decimal("0.00"))

        self.client.force_login(self.user)
        purchase_page = self.client.get(
            reverse("purchase_matches", kwargs={"period_id": period.id})
        )
        self.assertContains(purchase_page, "期初暂估")
        sales_page = self.client.get(
            reverse("sales_uninvoiced", kwargs={"period_id": period.id})
        )
        self.assertContains(sales_page, "发出商品跨月滚动勾稽")
        manual_invoice = PurchaseInvoiceLine.objects.create(
            period=period,
            batch=purchase_invoice_batch,
            invoice_no="PI-OPEN-MANUAL",
            invoice_date=date(2026, 1, 25),
            invoice_period=date(2026, 1, 1),
            supplier_code="SUP-OPEN",
            supplier_name="期初供应商",
            material_code="RM-OPEN",
            material_name="期初原料",
            quantity=Decimal("10"),
            remaining_quantity=Decimal("10"),
            unit_price_excl_tax=Decimal("11"),
            amount_excl_tax=Decimal("110"),
            match_mode=MatchingMode.MANUAL,
        )
        detail_url = reverse(
            "source_document_detail",
            kwargs={
                "period_id": period.id,
                "source_type": "purchase_invoice",
                "source_id": manual_invoice.id,
            },
        )
        detail_page = self.client.get(detail_url)
        self.assertContains(detail_page, "OPEN-GR-001")
        response = self.client.post(
            detail_url,
            {
                "action": "add_opening_match",
                "opening_id": purchase_opening.id,
                "matched_quantity": "10",
            },
        )
        self.assertEqual(response.status_code, 302)
        self.assertTrue(
            MatchAllocation.objects.filter(
                opening_purchase_accrual=purchase_opening,
                purchase_invoice_line=manual_invoice,
                is_manual=True,
            ).exists()
        )
        purchase_opening.refresh_from_db()
        self.assertEqual(purchase_opening.remaining_quantity, Decimal("50"))

    def test_user_controlled_workshop_sequence_supports_independent_settlement(self):
        period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        inventory_batch = self.create_batch(
            period, ImportBatch.Kind.INVENTORY, "inventory-chain"
        )
        contract_batch = self.create_batch(
            period, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-chain"
        )
        order_batch = self.create_batch(
            period, ImportBatch.Kind.WORK_ORDER, "order-chain"
        )
        raw = Material.objects.create(code="RAW", name="原料", uom="kg")
        semi = Material.objects.create(code="SEMI", name="半成品", uom="kg")
        final = Material.objects.create(code="FINAL", name="产成品", uom="kg")
        PurchaseContractLine.objects.create(
            period=period,
            batch=contract_batch,
            contract_no="PC-CHAIN",
            supplier_code="SUP",
            material_code=raw.code,
            quantity=Decimal("100"),
            remaining_quantity=Decimal("100"),
            unit_price_excl_tax=Decimal("10"),
        )
        order_a = WorkOrder.objects.create(
            period=period,
            batch=order_batch,
            order_no="WO-UP",
            production_dept_code="上游车间",
            plan_quantity=Decimal("10"),
        )
        order_b = WorkOrder.objects.create(
            period=period,
            batch=order_batch,
            order_no="WO-DOWN",
            production_dept_code="下游车间",
            plan_quantity=Decimal("5"),
        )
        MaterialIssue.objects.create(
            period=period,
            batch=inventory_batch,
            work_order_no=order_a.order_no,
            material_code=raw.code,
            material_name=raw.name,
            issue_quantity=Decimal("40"),
            actual_issue_quantity=Decimal("40"),
            issue_uom="kg",
        )
        MaterialIssue.objects.create(
            period=period,
            batch=inventory_batch,
            work_order_no=order_b.order_no,
            material_code=semi.code,
            material_name=semi.name,
            issue_quantity=Decimal("5"),
            actual_issue_quantity=Decimal("5"),
            issue_uom="kg",
        )
        WorkReport.objects.create(
            period=period,
            batch=inventory_batch,
            work_order_no=order_a.order_no,
            material_code=semi.code,
            material_name=semi.name,
            quantity=Decimal("10"),
        )
        WorkReport.objects.create(
            period=period,
            batch=inventory_batch,
            work_order_no=order_b.order_no,
            material_code=final.code,
            material_name=final.name,
            quantity=Decimal("5"),
        )
        self.create_txn(
            period,
            inventory_batch,
            txn_date=date(2026, 7, 1),
            document_no="GR-RAW",
            business_type="采购入库",
            direction=InventoryTransaction.Direction.IN,
            material=raw,
            quantity=Decimal("100"),
            supplier_code="SUP",
            contract_no="PC-CHAIN",
        )
        self.create_txn(
            period,
            inventory_batch,
            txn_date=date(2026, 7, 2),
            document_no="MI-RAW",
            business_type="投料",
            direction=InventoryTransaction.Direction.OUT,
            material=raw,
            quantity=Decimal("40"),
            order_no=order_a.order_no,
        )
        semi_receipt = self.create_txn(
            period,
            inventory_batch,
            txn_date=date(2026, 7, 3),
            document_no=order_a.order_no,
            business_type="报工",
            direction=InventoryTransaction.Direction.IN,
            material=semi,
            quantity=Decimal("10"),
            order_no=order_a.order_no,
        )
        self.create_txn(
            period,
            inventory_batch,
            txn_date=date(2026, 7, 4),
            document_no="MI-SEMI",
            business_type="投料",
            direction=InventoryTransaction.Direction.OUT,
            material=semi,
            quantity=Decimal("5"),
            order_no=order_b.order_no,
        )
        final_receipt = self.create_txn(
            period,
            inventory_batch,
            txn_date=date(2026, 7, 5),
            document_no=order_b.order_no,
            business_type="报工",
            direction=InventoryTransaction.Direction.IN,
            material=final,
            quantity=Decimal("5"),
            order_no=order_b.order_no,
        )
        final_issue = self.create_txn(
            period,
            inventory_batch,
            txn_date=date(2026, 7, 6),
            document_no="SO-FINAL",
            business_type="销售出库",
            direction=InventoryTransaction.Direction.OUT,
            material=final,
            quantity=Decimal("1"),
            customer_code="CUS",
        )

        settlements = {
            item.workshop: item
            for item in ensure_workshop_settlements(period)
        }
        update_workshop_sequences(
            period,
            {
                settlements["上游车间"].id: 1,
                settlements["下游车间"].id: 2,
            },
        )
        with self.assertRaisesMessage(ValueError, "请先完成前序车间"):
            settle_workshop(
                period,
                settlements["下游车间"].id,
                self.user,
            )
        upper_run = settle_workshop(
            period,
            settlements["上游车间"].id,
            self.user,
        )
        self.assertTrue(upper_run.converged)
        semi_receipt.refresh_from_db()
        final_receipt.refresh_from_db()
        self.assertEqual(semi_receipt.amount, Decimal("400.00"))
        self.assertIsNone(final_receipt.amount)
        self.assertEqual(
            WorkshopSettlement.objects.get(
                period=period, workshop="下游车间"
            ).status,
            WorkshopSettlement.Status.PENDING,
        )

        lower_run = settle_workshop(
            period,
            settlements["下游车间"].id,
            self.user,
        )
        self.assertTrue(lower_run.converged)
        semi_receipt.refresh_from_db()
        final_receipt.refresh_from_db()
        final_issue.refresh_from_db()
        self.assertEqual(semi_receipt.amount, Decimal("400.00"))
        self.assertEqual(final_receipt.amount, Decimal("200.00"))
        self.assertEqual(final_issue.amount, Decimal("40.00"))
        self.assertFalse(
            period.reconciliations.filter(
                status=ReconciliationResult.Status.FAIL
            ).exists()
        )
        period.refresh_from_db()
        self.assertEqual(period.status, AccountingPeriod.Status.CALCULATED)

    def test_locked_period_rolls_inventory_wip_order_and_contract_forward(self):
        june = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        inventory_batch = self.create_batch(
            june, ImportBatch.Kind.INVENTORY, "roll-inventory"
        )
        contract_batch = self.create_batch(
            june, ImportBatch.Kind.PURCHASE_CONTRACT, "roll-contract"
        )
        order_batch = self.create_batch(
            june, ImportBatch.Kind.WORK_ORDER, "roll-order"
        )
        raw = Material.objects.create(
            code="ROLL-RM", name="跨期原料", uom="kg"
        )
        PurchaseContractLine.objects.create(
            period=june,
            batch=contract_batch,
            contract_no="ROLL-PC",
            supplier_code="ROLL-SUP",
            material_code=raw.code,
            material_name=raw.name,
            quantity=Decimal("100"),
            remaining_quantity=Decimal("100"),
            unit_price_excl_tax=Decimal("2"),
        )
        self.create_txn(
            june,
            inventory_batch,
            txn_date=date(2026, 6, 1),
            document_no="OPEN-RM",
            business_type="期初结存",
            direction=InventoryTransaction.Direction.IN,
            material=raw,
            quantity=Decimal("10"),
        )
        opening_txn = june.inventory_transactions.get(document_no="OPEN-RM")
        opening_txn.unit_price = Decimal("10")
        opening_txn.amount = Decimal("100")
        opening_txn.price_source = "原始金额"
        opening_txn.save(update_fields=["unit_price", "amount", "price_source"])
        order = WorkOrder.objects.create(
            period=june,
            batch=order_batch,
            order_no="ROLL-WO",
            production_dept_code="跨期车间",
            plan_quantity=Decimal("10"),
        )
        MaterialIssue.objects.create(
            period=june,
            batch=inventory_batch,
            work_order_no=order.order_no,
            material_code=raw.code,
            material_name=raw.name,
            issue_quantity=Decimal("4"),
            actual_issue_quantity=Decimal("4"),
            issue_uom="kg",
        )
        self.create_txn(
            june,
            inventory_batch,
            txn_date=date(2026, 6, 15),
            document_no="MI-ROLL",
            business_type="投料",
            direction=InventoryTransaction.Direction.OUT,
            material=raw,
            quantity=Decimal("4"),
            order_no=order.order_no,
        )
        june_run = run_settlement(june, self.user)
        self.assertTrue(june_run.converged)
        june.status = AccountingPeriod.Status.LOCKED
        june.locked_by = self.user
        june.save(update_fields=["status", "locked_by"])

        july = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        # 模拟用户重传生产工单主档（含计划数量）
        WorkOrder.objects.create(
            period=july,
            order_no=order.order_no,
            production_dept_code="跨期车间",
            plan_quantity=Decimal("10"),
        )
        july_run = run_settlement(july, self.user)
        self.assertTrue(july_run.converged)
        opening_inventory = OpeningInventoryBalance.objects.get(
            period=july, material_code=raw.code
        )
        self.assertEqual(opening_inventory.source_period, june)
        self.assertEqual(opening_inventory.quantity, Decimal("6"))
        self.assertEqual(opening_inventory.amount, Decimal("60.00"))
        valuation = july.inventory_valuations.get(material_code=raw.code)
        self.assertEqual(valuation.opening_quantity, Decimal("6"))
        self.assertEqual(valuation.ending_amount, Decimal("60.00"))

        opening_wip = OpeningWip.objects.get(
            period=july, order_no=order.order_no
        )
        self.assertTrue(opening_wip.is_system_generated)
        self.assertEqual(opening_wip.source_period, june)
        self.assertEqual(opening_wip.total, Decimal("40.00"))
        carried_result = WorkOrderCostResult.objects.get(
            period=july, work_order__order_no=order.order_no
        )
        self.assertEqual(carried_result.wip_cost, Decimal("40.00"))
        self.assertEqual(carried_result.finished_cost, Decimal("0.00"))
        carried_contract = july.purchase_contracts.get(
            contract_no="ROLL-PC"
        )
        self.assertTrue(carried_contract.is_system_generated)
        self.assertEqual(carried_contract.quantity, Decimal("100"))
        self.assertFalse(
            july.reconciliations.filter(
                status=ReconciliationResult.Status.FAIL
            ).exists()
        )

    def test_previous_period_must_be_locked_before_next_period_settlement(self):
        AccountingPeriod.objects.create(month=date(2026, 6, 1))
        july = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        with self.assertRaisesMessage(ValueError, "尚未锁账"):
            run_settlement(july, self.user)

    def test_incomplete_order_outputs_finished_and_wip_costs(self):
        period = AccountingPeriod.objects.create(month=date(2026, 9, 1))
        inventory_batch = self.create_batch(
            period, ImportBatch.Kind.INVENTORY, "inventory-wip"
        )
        contract_batch = self.create_batch(
            period, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-wip"
        )
        order_batch = self.create_batch(
            period, ImportBatch.Kind.WORK_ORDER, "order-wip"
        )
        cost_batch = self.create_batch(
            period, ImportBatch.Kind.WORKSHOP_COST, "cost-wip"
        )
        raw = Material.objects.create(code="RAW-WIP", name="在制原料", uom="kg")
        product = Material.objects.create(code="FG-WIP", name="在制产品", uom="kg")
        PurchaseContractLine.objects.create(
            period=period,
            batch=contract_batch,
            contract_no="PC-WIP",
            supplier_code="SUP-WIP",
            material_code=raw.code,
            quantity=Decimal("100"),
            remaining_quantity=Decimal("100"),
            unit_price_excl_tax=Decimal("10"),
        )
        order = WorkOrder.objects.create(
            period=period,
            batch=order_batch,
            order_no="WO-WIP",
            production_dept_code="特纺",
            plan_quantity=Decimal("10"),
        )
        MaterialIssue.objects.create(
            period=period,
            batch=inventory_batch,
            work_order_no=order.order_no,
            material_code=raw.code,
            material_name=raw.name,
            issue_quantity=Decimal("50"),
            actual_issue_quantity=Decimal("50"),
            issue_uom="kg",
        )
        WorkReport.objects.create(
            period=period,
            batch=inventory_batch,
            work_order_no=order.order_no,
            material_code=product.code,
            material_name=product.name,
            quantity=Decimal("8"),
        )
        WorkshopCostPool.objects.create(
            period=period,
            batch=cost_batch,
            workshop="特纺",
            water_power=Decimal("100"),
            wages=Decimal("90"),
        )
        self.create_txn(
            period,
            inventory_batch,
            txn_date=date(2026, 9, 1),
            document_no="GR-WIP",
            business_type="采购入库",
            direction=InventoryTransaction.Direction.IN,
            material=raw,
            quantity=Decimal("100"),
            supplier_code="SUP-WIP",
            contract_no="PC-WIP",
        )
        self.create_txn(
            period,
            inventory_batch,
            txn_date=date(2026, 9, 2),
            document_no="MI-WIP",
            business_type="投料",
            direction=InventoryTransaction.Direction.OUT,
            material=raw,
            quantity=Decimal("50"),
            order_no=order.order_no,
            workshop="特纺",
        )
        receipt = self.create_txn(
            period,
            inventory_batch,
            txn_date=date(2026, 9, 28),
            document_no=order.order_no,
            business_type="报工",
            direction=InventoryTransaction.Direction.IN,
            material=product,
            quantity=Decimal("8"),
            order_no=order.order_no,
            workshop="特纺",
        )

        run = run_settlement(period)
        self.assertTrue(run.converged)
        result = order.cost_result
        self.assertEqual(result.total_cost, Decimal("690.00"))
        self.assertEqual(result.finished_material, Decimal("400.00"))
        self.assertEqual(result.finished_overhead, Decimal("80.00"))
        self.assertEqual(result.finished_wages, Decimal("80.00"))
        self.assertEqual(result.finished_cost, Decimal("560.00"))
        self.assertEqual(result.wip_cost, Decimal("130.00"))
        receipt.refresh_from_db()
        self.assertEqual(receipt.amount, Decimal("560.00"))
        self.assertEqual(receipt.unit_price, Decimal("70.00000000"))
        self.assertFalse(
            period.reconciliations.filter(
                status=ReconciliationResult.Status.FAIL
            ).exists()
        )


class ImportAndWebTests(SettlementFixtureMixin, TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="web-user", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 8, 1))

    def test_work_order_cost_filters_and_pulp_workshop(self):
        """工单成本页：期间/车间筛选生效；浆粕车间按上传流水补一行一个成品。"""
        wo = WorkOrder.objects.create(
            period=self.period, order_no="CS-F", production_dept_code="tefang",
            plan_quantity=100,
        )
        WorkOrderCostResult.objects.create(
            period=self.period, work_order=wo, product_code="P-F",
            product_name="纱线", plan_quantity=100, finished_quantity=100,
            wip_quantity=0, direct_material=200, finished_cost=200,
            wip_cost=0, total_cost=200,
        )
        # 浆粕：上传流水（原料领料 + 成品入库）与车间费用
        pulp_batch = self.create_batch(
            self.period, ImportBatch.Kind.PULP_FLOW, "woc-pulp"
        )
        product = Material.objects.create(
            code="21.03", name="芳纶浆粕DFP-103", uom="kg"
        )
        txn = self.create_txn(
            self.period, pulp_batch, txn_date=date(2026, 8, 20),
            document_no="PULP-OUT", business_type="生产入库",
            direction=InventoryTransaction.Direction.IN, material=product,
            quantity=Decimal("10"), order_no="PULP202608",
        )
        txn.amount = Decimal("0")
        txn.save(update_fields=["amount"])
        PulpMaterial.objects.create(
            period=self.period, batch=pulp_batch, txn_date=date(2026, 8, 5),
            document_no="PULP-IN", material_code="R52.101",
            material_name="对位芳纶短纤维", quantity=Decimal("8"),
            unit_price=Decimal("37.5"), amount=Decimal("300"),
            business_type="领料出库", uom="kg",
        )
        WorkshopCostPool.objects.create(
            period=self.period, batch=pulp_batch, workshop="浆粕车间",
            water_power=Decimal("100"), wages=Decimal("100"),
        )
        self.client.force_login(self.user)
        url = reverse("work_order_costs", args=[self.period.id])
        # 车间筛选：只看浆粕 → 单表里只有浆粕成品行
        resp = self.client.get(url, {"w_workshop": "浆粕车间"})
        rows = resp.context["rows"]
        self.assertEqual(len(rows), 1)
        row = rows[0]
        self.assertEqual(row.product_code, "21.03")
        self.assertEqual(row.wo_category, "浆粕")
        self.assertEqual(Decimal(row.finished_quantity), Decimal("10"))
        self.assertEqual(Decimal(row.direct_material), Decimal("300.00"))
        self.assertEqual(Decimal(row.water_power), Decimal("100.00"))
        self.assertEqual(Decimal(row.total_cost), Decimal("500.00"))
        self.assertEqual(Decimal(row.wip_cost), Decimal("0"))
        # 只看纺纱 → 单表里只剩纺纱那 1 行
        resp2 = self.client.get(url, {"w_workshop": "tefang"})
        self.assertEqual(len(resp2.context["rows"]), 1)
        self.assertEqual(
            resp2.context["rows"][0].work_order.production_dept_code, "tefang"
        )
        # 全部车间 → 纺纱 1 行 + 浆粕成品 1 行（本测试只造了 1 个浆粕成品）
        resp3 = self.client.get(url)
        self.assertEqual(resp3.context["row_count"], 2)
        self.assertEqual(resp3.context["subtotal"]["count"], 2)
        # 合计＝两行相加（400 + 200 的材料）
        self.assertEqual(
            resp3.context["subtotal"]["direct_material"], Decimal("500.00")
        )
        # 导出（全部车间）也带上浆粕成品行
        exp = self.client.get(
            reverse("export_listing", args=[self.period.id, "work_order_cost"])
        )
        wb = load_workbook(BytesIO(exp.content), read_only=True)
        try:
            ws = wb[wb.sheetnames[0]]
            data = list(ws.iter_rows(min_row=2, values_only=True))
        finally:
            wb.close()
        pulp_rows = [r for r in data if r[2] == "浆粕车间" and r[4] == "21.03"]
        self.assertEqual(len(pulp_rows), 1)
        self.assertEqual(pulp_rows[0][1], "浆粕车间（按上传流水）")

    def test_pagination_keeps_scroll_position(self):
        """翻页（上一页/下一页）不再跳回页顶：base.html 的脚本要接管 .pagination 链接。"""
        base = (Path(settings.BASE_DIR) / "templates" / "base.html").read_text(
            encoding="utf-8"
        )
        self.assertIn("closest(\".pagination a\")", base)
        self.assertIn("savePos(); return;", base)

    def test_work_order_costs_page_split_by_workshop(self):
        """工单成本页（用户 2026-09-21 二次调整）：只保留一张表，车间/期间用筛选切换。"""
        wo_tefang = WorkOrder.objects.create(
            period=self.period, order_no="CS-A",
            production_dept_code="tefang", plan_quantity=100,
        )
        wo_jizhi = WorkOrder.objects.create(
            period=self.period, order_no="JZ-001",
            production_dept_code="JiZhi", plan_quantity=50,
        )
        WorkOrderCostResult.objects.create(
            period=self.period, work_order=wo_tefang,
            product_code="P1", product_name="粗纱", plan_quantity=100,
            finished_quantity=80, wip_quantity=20,
            direct_material=1000, water_power=10, wages=20,
            finished_cost=900, wip_cost=130, total_cost=1030,
        )
        WorkOrderCostResult.objects.create(
            period=self.period, work_order=wo_jizhi,
            product_code="P2", product_name="坯布", plan_quantity=50,
            finished_quantity=50, wip_quantity=0,
            direct_material=500, steam=5, wages=10,
            finished_cost=520, wip_cost=0, total_cost=520,
        )
        self.client.force_login(self.user)
        resp = self.client.get(reverse("work_order_costs", args=[self.period.id]))
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        # 只保留一张表（用户 2026-09-21）：车间用筛选切换，不再每车间一张
        self.assertEqual(html.count("<table>"), 1)
        # 单表里有「车间」列（区分不同车间的行）
        self.assertIn("<th>车间</th>", html)
        self.assertIn("纺纱车间", html)
        self.assertIn("机织车间", html)
        # 单表末尾一个合计行
        self.assertEqual(html.count("合计（2 行）"), 1)
        # 车间导航按钮（页签）已按用户 2026-09-21 删除：改成「期间 + 车间」筛选
        self.assertNotIn("ws-tab", html)
        self.assertNotIn('name="ws-tab"', html)

    def test_uses_roving_shown_for_all_workshops(self):
        """用户 2026-09-30：「是否用粗纱」＝**BOM 100% 全是粗纱**才显示 是。

        混料单（粗纱＋其它料，如 RAJ031A-HF20260603）显示 否；所有车间都显示 是/否。
        """
        batch = self.create_batch(self.period, ImportBatch.Kind.WORK_ORDER, "rov-all")
        for order_no, dept, bom_codes in (
            ("ROV-JZ-1", "JiZhi", ("BC-16.0001CS",)),                 # 100% 粗纱 → 是
            ("ROV-JZ-2", "JiZhi", ()),                                # 不含粗纱 → 否
            ("ROV-RZ-1", "RanZheng", ("BC-16.0001CS",)),              # 100% 粗纱 → 是
            ("ROV-TF-1", "tefang", ()),                               # 无 BOM → 否
            ("ROV-MIX-1", "tefang", ("BC-16.0001CS", "R11.101")),     # 混料 → 否（用户重点）
        ):
            wo = WorkOrder.objects.create(
                period=self.period, batch=batch, order_no=order_no,
                production_dept_code=dept, plan_quantity=Decimal("100"),
            )
            WorkOrderCostResult.objects.create(
                period=self.period, work_order=wo, product_code="P-" + order_no,
                product_name="产品", plan_quantity=Decimal("100"),
                finished_quantity=Decimal("100"), direct_material=Decimal("100"),
                finished_cost=Decimal("100"), total_cost=Decimal("100"),
            )
            for index, code in enumerate(bom_codes):
                BomLine.objects.create(
                    period=self.period, batch=batch, work_order_no=order_no,
                    material_code="P-%d" % index, sub_item_code=code,
                    qty_numerator=Decimal("1"), qty_denominator=Decimal("1"),
                )
        self.client.force_login(self.user)
        resp = self.client.get(reverse("work_order_costs", args=[self.period.id]))
        got = {r.work_order.order_no: r.uses_roving for r in resp.context["rows"]}
        self.assertEqual(got["ROV-JZ-1"], "是")
        self.assertEqual(got["ROV-JZ-2"], "否")
        self.assertEqual(got["ROV-RZ-1"], "是")
        self.assertEqual(got["ROV-TF-1"], "否")
        self.assertEqual(got["ROV-MIX-1"], "否")   # 混料单不是"全粗纱"

    def test_unit_price_uses_bom_when_material_was_backfilled(self):
        """用户 2026-09-30：材料里含"按 BOM 反推补投"时，投料单价＝**BOM 折算价**（不是实领加权价）。

        实测 TDE025-BF20260720：只投了氨纶（25.66/kg），而材料 20.80 ÷ 报工 1.6 ＝ 13.00
        ＝ BOM 折算价 12.9991（含没投的涤纶 0.75×8.7788）。
        """
        from settlement.services.workorder_costing import (
            annotate_work_order_cost_rows, calculate_work_order_costs,
        )
        from settlement.models import InventoryValuation

        batch = self.create_batch(self.period, ImportBatch.Kind.WORK_ORDER, "unit-price")
        for code, price in (("R33.603", "8.7788"), ("R35.103", "25.66")):
            InventoryValuation.objects.create(
                period=self.period, material_code=code,
                weighted_average_price=Decimal(price),
            )
        WorkOrder.objects.create(
            period=self.period, batch=batch, order_no="UP-1",
            production_dept_code="BaoFu", plan_quantity=Decimal("1.6"),
        )
        # BOM：涤纶 0.75、氨纶 0.25 → Σ(单耗×价) ＝ 6.5841 ＋ 6.4150 ＝ 12.9991
        BomLine.objects.create(
            period=self.period, batch=batch, work_order_no="UP-1",
            material_code="P-A", sub_item_code="R33.603",
            qty_numerator=Decimal("75"), qty_denominator=Decimal("100"),
        )
        BomLine.objects.create(
            period=self.period, batch=batch, work_order_no="UP-1",
            material_code="P-B", sub_item_code="R35.103",
            qty_numerator=Decimal("25"), qty_denominator=Decimal("100"),
        )
        # 只投氨纶 0.4（涤纶没投）→ 情形③ 差额反推把涤纶 1.2 kg 的料补进来
        MaterialIssue.objects.create(
            period=self.period, batch=batch, work_order_no="UP-1",
            material_code="R35.103", material_name="氨纶长丝",
            material_category="长丝", issue_quantity=Decimal("0.4"),
            returned_quantity=Decimal("0"), actual_issue_quantity=Decimal("0.4"),
            issue_uom="千克",
        )
        WorkReport.objects.create(
            period=self.period, batch=batch,
            report_time=datetime(2026, 7, 3, 9, 0),
            work_order_no="UP-1", quantity=Decimal("1.6"),
        )
        prices = {"R33.603": Decimal("8.7788"), "R35.103": Decimal("25.66")}
        results = calculate_work_order_costs(self.period, None, prices)
        item = next(r for r in results if r.work_order.order_no == "UP-1")
        # 材料 ＝ 实领 10.26 ＋ 按 BOM 补投 10.54 ＝ 20.80
        self.assertEqual(item.direct_material, Decimal("20.80"))
        annotate_work_order_cost_rows(self.period, [item])
        # 投料单价 ＝ BOM 折算价 12.9991（不是实领氨纶的 25.66）
        self.assertEqual(item.current_unit_price, Decimal("12.9991"))
        self.assertTrue(item.current_unit_price_from_bom)

    def test_work_order_cost_detail_has_order_period_column(self):
        """工单成本明细（用户 2026-09-27）：原来的「期间」改「核算期间」，并加一列「工单期间」。"""
        from settlement.models import WorkOrderDetail

        batch = self.create_batch(self.period, ImportBatch.Kind.WORK_ORDER, "op")
        wo = WorkOrder.objects.create(
            period=self.period, batch=batch,
            order_no="OP-1", production_dept_code="tefang", plan_quantity=100,
        )
        WorkOrderCostResult.objects.create(
            period=self.period, work_order=wo, product_code="P9",
            product_name="纱线", plan_quantity=100, finished_quantity=100,
            direct_material=100, finished_cost=100, total_cost=100,
        )
        WorkOrderDetail.objects.create(
            period=self.period, batch=batch,
            order_no="OP-1", created_at="2026-05-09",
        )
        self.client.force_login(self.user)
        resp = self.client.get(reverse("work_order_costs", args=[self.period.id]))
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        heads = re.findall(r"<th[^>]*>(.*?)</th>", html, re.S)
        heads = [re.sub(r"<[^>]+>", "", h).strip() for h in heads]
        self.assertEqual(heads[:3], ["核算期间", "工单期间", "车间"])
        # 工单期间＝下单月份（2026-05），核算期间＝本期（2026-07）
        row = next(r for r in resp.context["rows"] if r.work_order.order_no == "OP-1")
        self.assertEqual(row.order_period_label, "2026-05")
        self.assertEqual(resp.context["period_label"], "2026-08")
        self.assertIn("2026-05", html)
        self.assertIn('name="w_period"', html)
        self.assertIn('name="w_workshop"', html)
        # 工单编号筛选（用户 2026-09-30）：明细表那一排要有输入框，能按工单号（部分匹配）筛
        form = html.split('<form class="woc-filter-row"', 1)[1].split("</form>", 1)[0]
        self.assertIn("工单编号", form)
        self.assertIn('name="order_no"', form)
        url = reverse("work_order_costs", args=[self.period.id])
        filtered = self.client.get(url, {"order_no": "OP"})
        numbers = [r.work_order.order_no for r in filtered.context["rows"]]
        self.assertIn("OP-1", numbers)
        self.assertTrue(all("OP" in n for n in numbers), numbers)
        filtered = self.client.get(url, {"order_no": "不存在的单号"})
        self.assertEqual(filtered.context["rows"], [])
        # 输入框回填（筛选后仍显示原值）
        self.assertIn('value="OP-1"', self.client.get(url, {"order_no": "OP-1"}).content.decode("utf-8"))
        # 车间下拉含两个车间（+ 全部车间 + 浆粕车间）
        self.assertIn('<option value="tefang"', html)
        self.assertIn('<option value="JiZhi"', html)
        self.assertIn("浆粕车间", html)
        # 类别/计划完工/实际完工/完成状态 四列存在
        self.assertIn("<th>类别</th>", html)
        self.assertIn("<th>计划完工时间</th>", html)
        self.assertIn("<th>实际完工时间</th>", html)
        self.assertIn("<th>完成状态</th>", html)
        # 导出Excel 就在筛选行里（与筛选/清空、打印本表同一排）
        form = html.split('<form class="woc-filter-row"', 1)[1].split("</form>", 1)[0]
        for token in (">筛选<", ">清空<", "导出Excel", 'class="print-here"'):
            self.assertIn(token, form, token)
        self.assertLess(form.index(">清空<"), form.index("导出Excel"))
        self.assertLess(form.index("导出Excel"), form.index('class="print-here"'))
        # 页头不再有导出按钮
        self.assertNotIn("导出Excel", html.split("</section>", 1)[0])
        # 用户 2026-09-21：删掉「工单总成本 / 完工成本 / 在产品 / 直接材料」那个小模块
        self.assertNotIn("工单总成本</span>", html)
        self.assertNotIn('class="summary-strip"', html)
        # 纺纱车间无天然气：天然气列只出现在非纺纱车间表（本测试仅机织有，故全页 1 处）
        self.assertEqual(html.count("<th class=\"num\">天然气</th>"), 1)
        self.assertEqual(html.count("<th class=\"num\">天然气分配率</th>"), 1)
        # 单独导出纺纱：Excel 无天然气列；单独导出机织：有
        resp = self.client.get(
            reverse("export_listing", args=[self.period.id, "work_order_cost"])
            + "?workshop=tefang"
        )
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(BytesIO(resp.content))
        header_row = [c.value for c in wb.active[1]]
        self.assertNotIn("天然气", header_row)
        resp = self.client.get(
            reverse("export_listing", args=[self.period.id, "work_order_cost"])
            + "?workshop=JiZhi"
        )
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(BytesIO(resp.content))
        header_row = [c.value for c in wb.active[1]]
        self.assertIn("天然气", header_row)
        # 本期投料/本期完工/本期完工率 在 累计投料/累计完工/累计完工率 之前
        self.assertIn("本期投料", header_row)
        self.assertIn("本期完工", header_row)
        self.assertIn("本期完工率", header_row)
        self.assertLess(
            header_row.index("本期投料"), header_row.index("累计投料")
        )
        self.assertLess(
            header_row.index("本期完工"), header_row.index("累计完工")
        )
        self.assertLess(
            header_row.index("本期完工率"), header_row.index("累计完工率")
        )

    def test_historical_price_import_and_price_table_override(self):
        """历史单价：导入后写入 HistoricalPrice、单价表页面/导出优先显示上传价；
        且导入历史单价不清理本期结算输出（库存计价）。"""
        from decimal import Decimal as D

        from settlement.models import HistoricalPrice, InventoryValuation

        # 系统计价：11.01 入库 10 件/1000 元 -> 系统入库单价 100
        InventoryValuation.objects.create(
            period=self.period, material_code="11.01",
            material_name="示例纱线", receipt_quantity=10, receipt_amount=1000,
            issue_quantity=5, issue_amount=600,
        )
        self.client.force_login(self.user)
        # 上传历史单价：覆盖 11.01 入库单价 88.88、出库单价 77.77
        csv_content = (
            "期间,物料编码,物料名称,入库单价,出库单价\n"
            "2026-08,11.01,示例纱线,88.88,77.77\n"
            "2026-08,H11.160,硝酸钠,5.55,\n"
        ).encode("utf-8-sig")
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                batch = create_import_batch(
                    period=self.period,
                    kind=ImportBatch.Kind.HISTORICAL_PRICE,
                    uploaded_file=SimpleUploadedFile(
                        "历史单价.csv", csv_content, "text/csv"
                    ),
                    user=self.user,
                )
                process_batch(batch)
        self.assertEqual(batch.row_count, 2)
        self.assertEqual(
            HistoricalPrice.objects.filter(period=self.period).count(), 2
        )
        hp = HistoricalPrice.objects.get(period=self.period, material_code="11.01")
        self.assertEqual(hp.receipt_price, D("88.88000000"))
        self.assertEqual(hp.issue_price, D("77.77000000"))
        # 导入历史单价不清理结算输出：库存计价仍在
        self.assertTrue(
            InventoryValuation.objects.filter(period=self.period).exists()
        )
        # 单价表页面显示上传价（覆盖系统 100/120）
        resp = self.client.get(reverse("price_table", args=[self.period.id]))
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        self.assertIn("88.8800", html)
        self.assertIn("77.7700", html)
        self.assertNotIn("100.0000", html)
        # 导出含上传价：两个工作表（入库单价 / 出库单价），列为 序号/品类/品名编码/品名名称 + 当年各月
        resp = self.client.get(
            reverse("price_table", args=[self.period.id]) + "?export=1"
        )
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(BytesIO(resp.content))
        self.assertEqual(wb.sheetnames, ["入库单价", "出库单价"])
        month_label = f"{self.period.month.year}年{self.period.month.month}月"
        for sheet_name, expected in (("入库单价", 88.88), ("出库单价", 77.77)):
            ws = wb[sheet_name]
            headers = [
                ws.cell(2, col).value for col in range(1, ws.max_column + 1)
            ]
            self.assertEqual(headers[:4], ["序号", "品类", "品名编码", "品名名称"])
            self.assertIn(month_label, headers)
            price_col = headers.index(month_label) + 1
            found = False
            for row in range(3, ws.max_row + 1):
                if ws.cell(row, 3).value == "11.01":
                    found = True
                    self.assertEqual(float(ws.cell(row, price_col).value), expected)
                    break
            self.assertTrue(found, f"{sheet_name} 未找到 11.01")

    def test_autofill_historical_prices_for_new_materials(self):
        """历史单价自动补全：**已锁账**期间内新增物料（无上传价）访问单价表时自动按系统计算价补入；
        已上传/已补全的物料不被覆盖；未锁账期间不补（避免把结算中途的 0.13 这种假价冻住）。"""
        from decimal import Decimal as D

        from settlement.models import HistoricalPrice, InventoryValuation

        # 物料A：有系统计价（入库 10/1000、出库 5/600），无上传历史单价 -> 应自动补
        InventoryValuation.objects.create(
            period=self.period, material_code="NEW-A",
            material_name="新原料A", receipt_quantity=10, receipt_amount=1000,
            issue_quantity=5, issue_amount=600, weighted_average_price=100,
        )
        # 物料B：金额数量全 0，无价格依据 -> 不补
        InventoryValuation.objects.create(
            period=self.period, material_code="NEW-B",
            material_name="无价物料B",
        )
        # 物料C：用户已上传 -> 不被覆盖
        HistoricalPrice.objects.create(
            period=self.period, material_code="NEW-C",
            material_name="已上传C", receipt_price=D("9.99000000"),
            issue_price=D("8.88000000"),
        )
        InventoryValuation.objects.create(
            period=self.period, material_code="NEW-C",
            material_name="已上传C", receipt_quantity=10, receipt_amount=5000,
            issue_quantity=5, issue_amount=3000, weighted_average_price=500,
        )

        self.client.force_login(self.user)
        # 未锁账：访问一次不补价（用户 2026-10-21 口径）
        resp = self.client.get(reverse("price_table", args=[self.period.id]))
        self.assertEqual(resp.status_code, 200)
        self.assertFalse(
            HistoricalPrice.objects.filter(period=self.period, material_code="NEW-A").exists()
        )

        # 锁账后再访问才补
        self.period.status = "locked"
        self.period.save(update_fields=["status"])
        resp = self.client.get(reverse("price_table", args=[self.period.id]))
        self.assertEqual(resp.status_code, 200)

        # 自动补全 NEW-A：入库价 1000/10=100、出库价 600/5=120
        hp_a = HistoricalPrice.objects.get(period=self.period, material_code="NEW-A")
        self.assertEqual(hp_a.receipt_price, D("100.0000"))
        self.assertEqual(hp_a.issue_price, D("120.0000"))
        # NEW-B 无价格依据不补
        self.assertFalse(
            HistoricalPrice.objects.filter(period=self.period, material_code="NEW-B").exists()
        )
        # NEW-C 已上传不被覆盖
        hp_c = HistoricalPrice.objects.get(period=self.period, material_code="NEW-C")
        self.assertEqual(hp_c.receipt_price, D("9.99000000"))
        self.assertEqual(hp_c.issue_price, D("8.88000000"))
        # 幂等：再次访问不新增
        self.client.get(reverse("price_table", args=[self.period.id]))
        self.assertEqual(
            HistoricalPrice.objects.filter(period=self.period).count(), 2
        )

    def test_price_table_shows_uploaded_price_without_settlement(self):
        """未结算期间（无库存计价）时，上传的历史单价也能在单价表显示。"""
        from decimal import Decimal as D

        from settlement.models import HistoricalPrice

        HistoricalPrice.objects.create(
            period=self.period, material_code="ONLY-HP",
            material_name="仅上传价物料",
            receipt_price=D("12.34000000"), issue_price=D("11.11000000"),
        )
        self.client.force_login(self.user)
        resp = self.client.get(reverse("price_table", args=[self.period.id]))
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        # 即使该期间没有任何库存计价，上传的物料也显示（含单价）
        self.assertIn("ONLY-HP", html)
        self.assertIn("12.3400", html)
        self.assertIn("11.1100", html)
        # 导出：该期间没有任何入库/出库，按"有进出才填月份"的口径两张表均为空表
        resp = self.client.get(
            reverse("price_table", args=[self.period.id]) + "?export=1"
        )
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(BytesIO(resp.content))
        self.assertEqual(wb.sheetnames, ["入库单价", "出库单价"])
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            self.assertEqual(ws.max_row, 2)  # 仅标题行 + 表头行

    def test_export_center_includes_price_table(self):
        """导出中心包含"单价表"，勾选后 ZIP 内可导出跨月单价表。"""
        from decimal import Decimal as D
        from zipfile import ZipFile

        from settlement.models import HistoricalPrice

        HistoricalPrice.objects.create(
            period=self.period, material_code="EXP-HP",
            material_name="导出测试物料",
            receipt_price=D("45.67000000"), issue_price=D("44.44000000"),
        )
        # 需要有实际进出，单价才会落到对应月份列（上传价覆盖系统价）
        from settlement.models import InventoryValuation

        InventoryValuation.objects.create(
            period=self.period, material_code="EXP-HP",
            material_name="导出测试物料",
            receipt_quantity=10, receipt_amount=1000,
            issue_quantity=5, issue_amount=600,
        )
        self.client.force_login(self.user)
        # 导出中心页面含单价表选项
        resp = self.client.get(reverse("export_center", args=[self.period.id]))
        self.assertEqual(resp.status_code, 200)
        self.assertContains(resp, "单价表（入库/出库·跨月对比）")
        # 勾选导出 ZIP
        resp = self.client.post(
            reverse("export_center", args=[self.period.id]),
            data={"reports": ["price_table"]},
        )
        self.assertEqual(resp.status_code, 200)
        self.assertIn("application/zip", resp["Content-Type"])
        zf = ZipFile(BytesIO(resp.content))
        names = zf.namelist()
        self.assertTrue(any("单价表" in n for n in names))
        wb = load_workbook(
            BytesIO(zf.read([n for n in names if "单价表" in n][0]))
        )
        self.assertEqual(wb.sheetnames, ["入库单价", "出库单价"])
        month_label = f"{self.period.month.year}年{self.period.month.month}月"
        for sheet_name, expected in (("入库单价", 45.67), ("出库单价", 44.44)):
            ws = wb[sheet_name]
            headers = [
                ws.cell(2, col).value for col in range(1, ws.max_column + 1)
            ]
            price_col = headers.index(month_label) + 1
            found = False
            for row in range(3, ws.max_row + 1):
                if ws.cell(row, 3).value == "EXP-HP":
                    found = True
                    self.assertEqual(float(ws.cell(row, price_col).value), expected)
                    break
            self.assertTrue(found, f"{sheet_name} 未找到 EXP-HP")

    def test_ledger_pages_hidden_and_exports_available(self):
        """工单/投料/报工台账不再出现在页面导航，仅在报表导出中心可导出（原样明细）。"""
        from zipfile import ZipFile

        from settlement.models import MaterialIssue, WorkReport

        WorkOrder.objects.create(
            period=self.period, order_no="LEDGER-1",
            order_type="普通工单", production_dept_code="tefang",
            status="已完工", plan_quantity=100,
        )
        MaterialIssue.objects.create(
            period=self.period, task_no="TASK-L", work_order_no="LEDGER-1",
            material_code="RM-L", material_name="示例原料",
            issue_quantity=50, actual_issue_quantity=50, issue_uom="kg",
        )
        WorkReport.objects.create(
            period=self.period, task_no="TASK-L", work_order_no="LEDGER-1",
            material_code="FG-L", material_name="示例成品",
            quantity=80, uom="kg",
        )
        self.client.force_login(self.user)
        # 页面导航不含三个台账链接
        for page_url_name in ("period_detail", "work_order_costs"):
            html = self.client.get(
                reverse(page_url_name, args=[self.period.id])
            ).content.decode("utf-8")
            for name in ("工单台账", "投料台账", "报工台账"):
                self.assertNotIn(name, html)
        # 导出中心含三项
        html = self.client.get(
            reverse("export_center", args=[self.period.id])
        ).content.decode("utf-8")
        for name in ("工单台账（原样明细）", "投料台账（原样明细）", "报工台账（原样明细）"):
            self.assertIn(name, html)
        # 勾选导出 ZIP 含三个台账且数据完整
        resp = self.client.post(
            reverse("export_center", args=[self.period.id]),
            data={"reports": [
                "work_order_ledger", "material_issue_ledger", "work_report_ledger",
            ]},
        )
        self.assertEqual(resp.status_code, 200)
        self.assertIn("application/zip", resp["Content-Type"])
        zf = ZipFile(BytesIO(resp.content))
        names = zf.namelist()
        for marker, col, code in (
            ("工单台账", 0, "LEDGER-1"),
            ("投料台账", 3, "RM-L"),
            ("报工台账", 4, "FG-L"),
        ):
            fname = [n for n in names if marker in n][0]
            ws = load_workbook(BytesIO(zf.read(fname))).active
            rows = list(ws.iter_rows(values_only=True))
            self.assertTrue(any(r and r[col] == code for r in rows), marker)

    def test_work_order_detail_report_export(self):
        """导出中心新增"工单明细"报表：9列模板（创建时间/车间/工单号/计划完工/实际完工/完成状态/计划数量/投料数量/入库数量）。"""
        from zipfile import ZipFile

        from settlement.models import MaterialIssue, WorkReport

        wo = WorkOrder.objects.create(
            period=self.period, order_no="DETAIL-1",
            production_dept_code="tefang", plan_quantity=100,
            planned_finish=None, created_at=None,
        )
        MaterialIssue.objects.create(
            period=self.period, work_order_no="DETAIL-1",
            material_code="RM-D", material_name="料",
            actual_issue_quantity=60, issue_uom="kg",
        )
        WorkReport.objects.create(
            period=self.period, work_order_no="DETAIL-1",
            material_code="FG-D", material_name="品",
            quantity=40, uom="kg", report_time=None,
        )
        self.client.force_login(self.user)
        html = self.client.get(
            reverse("export_center", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertIn("工单明细", html)
        resp = self.client.post(
            reverse("export_center", args=[self.period.id]),
            data={"reports": ["work_order_detail"]},
        )
        self.assertEqual(resp.status_code, 200)
        zf = ZipFile(BytesIO(resp.content))
        fname = [n for n in zf.namelist() if "工单明细" in n][0]
        ws = load_workbook(BytesIO(zf.read(fname))).active
        rows = list(ws.iter_rows(values_only=True))
        header = [str(c) for c in rows[0]]
        self.assertEqual(
            header,
            ["创建时间", "生产车间", "工单编号", "计划完工时间", "实际完工时间",
             "完成状态", "计划数量", "投料数量", "入库数量"],
        )
        found = [r for r in rows[1:] if r and r[2] == "DETAIL-1"]
        self.assertTrue(found)
        r = found[0]
        self.assertEqual(r[1], "纺纱车间")
        self.assertEqual(float(r[6]), 100.0)   # 计划数量
        self.assertEqual(float(r[7]), 60.0)    # 投料数量
        self.assertEqual(float(r[8]), 0.0)     # 入库数量（无报工入库流水）

    def test_work_order_detail_upload_and_original_export(self):
        """工单明细作为上传模板（第19类）：上传后入库，导出中心按上传原样导出。"""
        from decimal import Decimal as D
        from zipfile import ZipFile

        from openpyxl import Workbook

        from settlement.models import WorkOrderDetail

        wb = Workbook()
        ws = wb.active
        ws.title = "工单明细"
        ws.append(["工单明细"])
        ws.append(["说明"])
        ws.append(["创建时间", "生产车间", "工单编号", "计划完工时间", "实际完工时间",
                   "完成状态", "计划数量", "投料数量", "入库数量"])
        ws.append(["2026-07-02 08:00", "纺纱车间", "UP-D1", "2026-07-06 18:00",
                   "2026-07-06 17:00", "已完成", "200", "210", "198"])
        stream = BytesIO()
        wb.save(stream)
        stream.seek(0)
        self.client.force_login(self.user)
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                batch = create_import_batch(
                    period=self.period,
                    kind=ImportBatch.Kind.WORK_ORDER_DETAIL,
                    uploaded_file=SimpleUploadedFile(
                        "工单明细.xlsx", stream.read(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ),
                    user=self.user,
                )
                process_batch(batch)
                self.assertEqual(batch.row_count, 1)
                d = WorkOrderDetail.objects.get(period=self.period, order_no="UP-D1")
                self.assertEqual(d.workshop, "纺纱车间")
                self.assertEqual(d.plan_quantity, D("200"))
                self.assertEqual(d.issue_quantity, D("210"))
                self.assertEqual(d.receipt_quantity, D("198"))
                # 导出中心"工单明细"按上传原样
                resp = self.client.post(
                    reverse("export_center", args=[self.period.id]),
                    data={"reports": ["work_order_detail"]},
                )
                self.assertEqual(resp.status_code, 200)
                zf = ZipFile(BytesIO(resp.content))
                fname = [n for n in zf.namelist() if "工单明细" in n][0]
                wsx = load_workbook(BytesIO(zf.read(fname))).active
                rows = list(wsx.iter_rows(values_only=True))
                header = [str(c) for c in rows[0]]
                self.assertEqual(
                    header,
                    ["创建时间", "生产车间", "工单编号", "计划完工时间", "实际完工时间",
                     "完成状态", "计划数量", "投料数量", "入库数量"],
                )
                found = [r for r in rows[1:] if r and r[2] == "UP-D1"]
                self.assertTrue(found)
                self.assertEqual(found[0][1], "纺纱车间")
                self.assertEqual(float(found[0][6]), 200.0)

    def test_work_order_detail_export_keeps_upload_order(self):
        """工单明细导出按上传顺序（非字母序）。"""
        from zipfile import ZipFile

        from openpyxl import Workbook

        wb = Workbook()
        ws = wb.active
        ws.title = "工单明细"
        ws.append(["工单明细"])
        ws.append(["说明"])
        ws.append(["创建时间", "生产车间", "工单编号", "计划完工时间", "实际完工时间",
                   "完成状态", "计划数量", "投料数量", "入库数量"])
        for r in [
            ["2026-07-02 08:00", "纺纱车间", "Z-001", "", "", "已完成", "200", "210", "198"],
            ["2026-07-02 08:05", "染色车间", "A-002", "", "", "部分完成", "300", "280", "150"],
            ["2026-07-02 08:10", "机织车间", "M-003", "", "", "未开工", "100", "0", "0"],
        ]:
            ws.append(r)
        stream = BytesIO()
        wb.save(stream)
        stream.seek(0)
        self.client.force_login(self.user)
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                batch = create_import_batch(
                    period=self.period,
                    kind=ImportBatch.Kind.WORK_ORDER_DETAIL,
                    uploaded_file=SimpleUploadedFile(
                        "工单明细.xlsx", stream.read(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ),
                    user=self.user,
                )
                process_batch(batch)
                self.assertEqual(batch.row_count, 3)
                resp = self.client.post(
                    reverse("export_center", args=[self.period.id]),
                    data={"reports": ["work_order_detail"]},
                )
                self.assertEqual(resp.status_code, 200)
                zf = ZipFile(BytesIO(resp.content))
                fname = [n for n in zf.namelist() if "工单明细" in n][0]
                wsx = load_workbook(BytesIO(zf.read(fname))).active
                rows = list(wsx.iter_rows(values_only=True))
                orders = [r[2] for r in rows[1:] if r and r[2]]
                self.assertEqual(orders, ["Z-001", "A-002", "M-003"])

    def test_opening_wip_multisheet_template_import(self):
        """期初在产品模板改为多工作表（每车间一个 sheet），导入时合并全部 sheet。"""
        from decimal import Decimal as D

        from openpyxl import Workbook

        from settlement.models import OpeningWip

        wb = Workbook()
        ws1 = wb.active
        ws1.title = "染色车间"
        ws1.append(["期初在产品 · 染色车间"])
        ws1.append(["说明"])
        ws1.append(["工单号", "材料成本", "水电", "蒸汽", "天然气", "折旧", "机物料", "染料助剂", "工资"])
        ws1.append(["WIP-RZ-1", "1000", "50", "30", "20", "40", "10", "60", "80"])
        ws2 = wb.create_sheet("纺纱车间")
        ws2.append(["期初在产品 · 纺纱车间"])
        ws2.append(["说明"])
        ws2.append(["工单号", "材料成本", "水电", "蒸汽", "天然气", "折旧", "机物料", "染料助剂", "工资"])
        ws2.append(["WIP-TF-1", "2000", "80", "60", "", "50", "20", "", "120"])
        stream = BytesIO()
        wb.save(stream)
        stream.seek(0)

        self.client.force_login(self.user)
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                batch = create_import_batch(
                    period=self.period,
                    kind=ImportBatch.Kind.OPENING_WIP,
                    uploaded_file=SimpleUploadedFile(
                        "期初在产品.xlsx", stream.read(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ),
                    user=self.user,
                )
                process_batch(batch)
        self.assertEqual(batch.row_count, 2)
        self.assertEqual(batch.error_count, 0)
        rz = OpeningWip.objects.get(period=self.period, order_no="WIP-RZ-1")
        self.assertEqual(rz.material, D("1000.00"))
        self.assertEqual(rz.natural_gas, D("20.00"))
        tf = OpeningWip.objects.get(period=self.period, order_no="WIP-TF-1")
        self.assertEqual(tf.material, D("2000.00"))
        self.assertEqual(tf.natural_gas, D("0.00"))  # 纺纱留空天然气
        self.assertEqual(tf.dye_auxiliary, D("0.00"))

    def test_ledger_export_uses_uploaded_format(self):
        """台账导出格式 = 用户上传文件的列（原样），数据 = 数据库全部记录。"""
        from zipfile import ZipFile

        from openpyxl import Workbook

        from settlement.models import MaterialIssue

        # 上传一个投料记录 xlsx（自定义表头 4 列 + 2 行数据）
        wb = Workbook()
        ws = wb.active
        ws.title = "投料"
        ws.append(["工单编号", "物料编码", "投料数量", "备注"])
        ws.append(["WO-U1", "RM-U1", "50", "A"])
        ws.append(["WO-U2", "RM-U2", "80", "B"])
        stream = BytesIO()
        wb.save(stream)
        stream.seek(0)
        self.client.force_login(self.user)
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                batch = create_import_batch(
                    period=self.period,
                    kind=ImportBatch.Kind.MATERIAL_ISSUE,
                    uploaded_file=SimpleUploadedFile(
                        "投料记录.xlsx", stream.read(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ),
                    user=self.user,
                )
                process_batch(batch)
                # 导出投料台账：表头 = 上传表头（4列），数据 = 2 行
                resp = self.client.post(
                    reverse("export_center", args=[self.period.id]),
                    data={"reports": ["material_issue_ledger"]},
                )
                self.assertEqual(resp.status_code, 200)
                zf = ZipFile(BytesIO(resp.content))
                fname = [n for n in zf.namelist() if "投料台账" in n][0]
                ws = load_workbook(BytesIO(zf.read(fname))).active
                rows = list(ws.iter_rows(values_only=True))
                header = [str(c) for c in rows[0]]
                # 投料台账 = 上传表头（4列） + 追加"单价"列（v1.5.28）
                self.assertEqual(header, ["工单编号", "物料编码", "投料数量", "备注", "单价"])
                self.assertEqual(len(rows) - 1, 2)
                # 数据内容映射正确
                data = {r[0]: r for r in rows[1:]}
                self.assertIn("WO-U1", data)
                self.assertEqual(float(data["WO-U1"][2]), 50.0)
        # 上传的两条已入库（4列表头里：工单编号→order_no 等都在 ALIASES 中）
        self.assertEqual(
            MaterialIssue.objects.filter(period=self.period).count(), 2
        )


    def test_csv_import_and_authenticated_pages(self):
        content = (
            "业务日期,单据号,业务类型,方向,物料编码,物料名称,数量,单位\n"
            "2026-08-01,GR-1,采购入库,入库,RM-1,原料一,10,kg\n"
        ).encode("utf-8-sig")
        uploaded = SimpleUploadedFile("inventory.csv", content, "text/csv")
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                batch = create_import_batch(
                    period=self.period,
                    kind=ImportBatch.Kind.INVENTORY,
                    uploaded_file=uploaded,
                    user=self.user,
                )
                process_batch(batch)
        self.assertEqual(self.period.inventory_transactions.count(), 1)
        txn = self.period.inventory_transactions.get()
        self.assertEqual(txn.material_code, "RM-1")
        self.assertEqual(txn.direction, InventoryTransaction.Direction.IN)

        self.client.force_login(self.user)
        self.assertEqual(self.client.get(reverse("dashboard")).status_code, 200)
        self.assertEqual(
            self.client.get(
                reverse("period_detail", kwargs={"period_id": self.period.id})
            ).status_code,
            200,
        )
        self.assertEqual(
            self.client.get(
                reverse("purchase_matches", kwargs={"period_id": self.period.id})
            ).status_code,
            200,
        )
        upload_page = self.client.get(
            reverse("upload_file", kwargs={"period_id": self.period.id})
        )
        self.assertEqual(upload_page.status_code, 200)
        self.assertContains(upload_page, "上传前必须注意")
        self.assertContains(upload_page, "十八类资料字段说明")
        self.assertContains(upload_page, "“替换”会删除本月整类资料")
        self.assertContains(upload_page, "同车间费用必须先汇总")
        self.assertContains(upload_page, "销售订单号、客户和产品一致")
        for kind, _label in ImportBatch.Kind.choices:
            template_response = self.client.get(
                reverse("download_template", kwargs={"kind": kind})
            )
            self.assertEqual(template_response.status_code, 200)
            self.assertIn(
                "attachment",
                template_response.headers.get("Content-Disposition", ""),
            )

    def test_inventory_template_new_format_imports_with_inferred_direction(self):
        template_dir = Path(settings.BASE_DIR) / "sample_templates"
        content = (template_dir / "库存出入库流水.csv").read_bytes()
        uploaded = SimpleUploadedFile("库存出入库流水.csv", content, "text/csv")
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                batch = create_import_batch(
                    period=self.period,
                    kind=ImportBatch.Kind.INVENTORY,
                    uploaded_file=uploaded,
                    user=self.user,
                )
                process_batch(batch)
        self.assertEqual(self.period.inventory_transactions.count(), 4)
        receipt = self.period.inventory_transactions.get(document_no="GR-001")
        self.assertEqual(receipt.txn_date, date(2026, 6, 5))
        self.assertEqual(receipt.direction, InventoryTransaction.Direction.IN)
        self.assertEqual(receipt.material_code, "RM-001")
        self.assertEqual(receipt.category_name, "原材料")
        self.assertEqual(receipt.batch_no, "B20260601")
        self.assertEqual(receipt.uom, "kg")
        self.assertEqual(receipt.auxiliary_quantity, Decimal("200"))
        self.assertEqual(receipt.auxiliary_uom, "件")
        self.assertEqual(receipt.supplier_name, "示例供应商")
        self.assertEqual(receipt.warehouse, "原料仓")
        self.assertEqual(
            self.period.inventory_transactions.get(document_no="MI-001").direction,
            InventoryTransaction.Direction.OUT,
        )
        self.assertEqual(
            self.period.inventory_transactions.get(document_no="PR-001").direction,
            InventoryTransaction.Direction.IN,
        )
        self.assertEqual(
            self.period.inventory_transactions.get(document_no="SO-001").direction,
            InventoryTransaction.Direction.OUT,
        )

    def test_all_supplied_templates_import_and_close_end_to_end(self):
        template_map = [
            (ImportBatch.Kind.INVENTORY, "库存出入库流水.csv"),
            (ImportBatch.Kind.PURCHASE_CONTRACT, "采购合同.csv"),
            (ImportBatch.Kind.PURCHASE_INVOICE, "采购发票.csv"),
            (ImportBatch.Kind.SALES_INVOICE, "销售发票.csv"),
            (ImportBatch.Kind.WORKSHOP_COST, "车间费用.csv"),
            (ImportBatch.Kind.OPENING_WIP, "期初在产品.xlsx"),
            (ImportBatch.Kind.OPENING_INVENTORY, "期初存货.csv"),
            (ImportBatch.Kind.ENDING_INVENTORY, "期末存货.csv"),
            (
                ImportBatch.Kind.OPENING_PURCHASE_ACCRUAL,
                "期初采购暂估明细.csv",
            ),
            # 期初发出商品：sample_templates 里的 期初发出商品明细.csv 是旧 17 列格式，
            # 与现行"同销售台账"23 列口径不一致，这里用内嵌的现行格式数据
            (
                ImportBatch.Kind.OPENING_GOODS_ISSUED,
                "期初发出商品明细.csv",
            ),
            (ImportBatch.Kind.WORK_ORDER, "生产工单.csv"),
            (ImportBatch.Kind.MATERIAL_ISSUE, "投料记录.csv"),
            (ImportBatch.Kind.WORK_REPORT, "报工记录.csv"),
        ]
        # 库存流水用含自动匹配列的完整格式，保证采购/生产/销售全链路数字成立；
        # 生产环节使用新口径：投料出库 + 报工入库（单据号=工单编号）。
        inventory_csv = (
            "业务日期,单据号,行号,业务类型,方向,仓库,物料编码,物料名称,规格型号,数量,单位,工单号,车间,供应商编码,供应商名称,客户编码,客户名称,合同号,发票号,不含税单价,金额,备注\n"
            "2026-06-05,GR-001,1,采购入库,入库,原料仓,RM-001,示例原料,,100,kg,,,SUP-001,示例供应商,,,PC-001,,,,示例行\n"
            "2026-06-10,MI-001,1,投料,出库,原料仓,RM-001,示例原料,,50,kg,WO-001,染整,,,,,,,,,\n"
            "2026-06-20,WO-001,1,报工,入库,成品仓,FG-001,示例产品,,10,kg,WO-001,染整,,,,,,,,,\n"
            "2026-06-24,SO-001,1,销售出库,出库,成品仓,FG-001,示例产品,,4,kg,SO-001,,,,CUS-001,示例客户,,,,,\n"
        ).encode("utf-8-sig")
        # 三个生产台账模板按新口径提供与库存流水一致的工单数据（WO-001）
        work_order_csv = (
            "工单编号,工单类型,标识码,计划数量,计划开始时间,计划完工时间,优先级,产线设备,班次,生产部门编号,生产主管编号,计划部门编号,计划员编号,备注,指定用料,外部单据编号,状态,是否使用物料清单,是否使用工艺路线,工序计划类型,启用SOP,生产模式,产线编号,订单号,物料颜色\n"
            "WO-001,普通工单,WO-001,10,,,,,,染整,,,,,,,,,,,,,,,\n"
        ).encode("utf-8-sig")
        material_issue_csv = (
            "生产任务编号,工序编号,工序名称,物料编码,物料名称,物料分类,物料属性,物料规格,投料数量,已回撤数量,实际投料数量,投料单位,辅助单位数量,辅助单位,标识码,批次号,质量状态,投料仓位,被替代物料编号,被替代物料名称,投料时间,投料人员,执行人员,工单编号,主产出物料编号,主产出物料名称,销售订单编号,投料记录编号\n"
            "TASK-001,OP10,投料,RM-001,示例原料,,,,50,0,50,kg,,,,,,,,,2026-06-10 09:00,,,WO-001,FG-001,示例产品,,\n"
        ).encode("utf-8-sig")
        work_report_csv = (
            "报工时间,生产任务编号,工序编号,工序名称,报工物料编号,报工物料名称,生产人员,质量状态,不良原因,报工数量,报工单位,辅助数量,辅助单位,批次号,标识码,流转卡编号,工时,工时单位,备注,部门,设备,工单编号,工单物料编号,工单物料名称,订单编号,特纺/包覆-纸管颜色,纺纱-班别,染色纤维-含水率,染色纤维-批号流水,织布-序号,织布-机台,产出仓位,工单部门\n"
            "2026-06-20 16:00,TASK-001,OP10,报工,FG-001,示例产品,,,,10,kg,,,,,,,,,,,WO-001,RM-001,示例原料,,,,,,,,\n"
        ).encode("utf-8-sig")
        # 期初发出商品：现行「同销售台账」23 列格式（旧 17 列 CSV 已不适用）
        opening_goods_issued_csv = (
            "序号,出库期间,开票状态,开票期间,销售渠道,销售公司,订单单号,客户名称,销售大类,产品编码,产品名称,规格型号,计量单位,销售数量,销售成本单价,销售成本\n"
            "1,2026-06,未开票,,直销/终端,射阳兰邦,OG-001,示例客户,成品,FG-001,示例产品,,kg,2,50,\n"
        ).encode("utf-8-sig")
        template_dir = Path(settings.BASE_DIR) / "sample_templates"
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                for kind, filename in template_map:
                    if kind == ImportBatch.Kind.INVENTORY:
                        content = inventory_csv
                    elif kind == ImportBatch.Kind.WORK_ORDER:
                        content = work_order_csv
                    elif kind == ImportBatch.Kind.MATERIAL_ISSUE:
                        content = material_issue_csv
                    elif kind == ImportBatch.Kind.WORK_REPORT:
                        content = work_report_csv
                    elif kind == ImportBatch.Kind.OPENING_GOODS_ISSUED:
                        content = opening_goods_issued_csv
                    elif kind == ImportBatch.Kind.WORKSHOP_COST:
                        # 车间费用用内嵌数据（染整 水电100/人工50），
                        # 不依赖 sample_templates/车间费用.csv（该文件为用户模板，含预填车间+合计行）
                        content = (
                            "车间,人工费,折旧费,水电费,蒸汽费,燃气费,染料助剂费,机物料费,加工费,合计,备注\n"
                            "染整,50,0,100,0,0,0,0,0,,\n"
                        ).encode("utf-8-sig")
                    else:
                        content = (template_dir / filename).read_bytes()
                    uploaded = SimpleUploadedFile(filename, content, "text/csv")
                    batch = create_import_batch(
                        period=self.period,
                        kind=kind,
                        uploaded_file=uploaded,
                        user=self.user,
                    )
                    process_batch(batch)

                run = run_settlement(self.period)

        self.assertTrue(run.converged)
        self.assertEqual(self.period.import_batches.count(), 13)
        self.assertEqual(self.period.inventory_transactions.count(), 4)
        self.assertEqual(self.period.work_orders.count(), 1)
        self.assertEqual(self.period.work_order_cost_results.count(), 1)
        result = self.period.work_order_cost_results.get()
        self.assertEqual(result.finished_cost, Decimal("650.00"))
        self.assertEqual(result.wip_cost, Decimal("0.00"))
        # 期末存货以用户上传报表为准：本测试上传了期末模板(OP-RM 100kg/1000元)
        uploaded = self.period.ending_inventory_balances.filter(
            is_system_generated=False
        )
        self.assertTrue(uploaded.exists())
        self.assertEqual(uploaded.first().quantity, Decimal("100"))
        # 系统自动生成期末行：无人工期末的物料按 期初+入库−出库 生成
        sys_ending = self.period.ending_inventory_balances.filter(
            is_system_generated=True
        )
        self.assertTrue(sys_ending.exists())
        # 上传了人工期末的 OP-RM 不被系统覆盖（以人工为准）
        op_rm_manual = self.period.ending_inventory_balances.filter(
            is_system_generated=False, material_code="OP-RM"
        )
        self.assertTrue(op_rm_manual.exists())
        # 勾稽期末用上传值；上传期末与系统计算不一致时，库存勾稽应反映该差异
        inv_check = self.period.reconciliations.get(
            check_code="inventory_amount"
        )
        valuation = self.period.inventory_valuations.get(
            material_code="RM-001"
        )
        # right = 出库金额 + 上传期末金额(1000)；left = 期初+入库
        self.assertNotEqual(
            inv_check.difference, Decimal("0.00"),
            "上传期末(1000)与系统计算期末不一致时应产生勾稽差异",
        )
        self.assertNotEqual(
            inv_check.right_amount,
            valuation.issue_amount + valuation.ending_amount,
            "勾稽期末应使用上传的期末金额而非系统计算值",
        )
        summary = build_work_order_cost_summary(self.period)
        self.assertEqual(summary["totals"]["current_material"], Decimal("500.00"))
        self.assertEqual(summary["totals"]["current_expenses"], Decimal("150.00"))
        self.assertEqual(summary["totals"]["current_input"], Decimal("650.00"))
        self.assertEqual(summary["totals"]["finished_cost"], Decimal("650.00"))
        self.assertEqual(summary["totals"]["ending_wip"], Decimal("0.00"))
        self.assertEqual(summary["totals"]["difference"], Decimal("0.00"))
        self.assertTrue(summary["totals"]["all_calculated"])

        self.client.force_login(self.user)
        production_cost_page = self.client.get(
            reverse(
                "production_cost_summary",
                kwargs={"period_id": self.period.id},
            )
        )
        self.assertEqual(production_cost_page.status_code, 200)
        self.assertContains(production_cost_page, "生产成本")
        self.assertContains(production_cost_page, "染整")
        self.assertContains(production_cost_page, "650.00")
        workshop_page = self.client.get(
            reverse(
                "production_cost_summary",
                kwargs={"period_id": self.period.id},
            ),
            {"workshop": "染整"},
        )
        self.assertContains(workshop_page, "染整工单成本明细")
        self.assertContains(workshop_page, "WO-001")

    def test_work_order_issue_report_templates_import_into_ledgers(self):
        """生产工单/投料记录/报工记录三类模板可按 ERP 格式导入台账。"""
        template_dir = Path(settings.BASE_DIR) / "sample_templates"
        cases = [
            (ImportBatch.Kind.WORK_ORDER, "生产工单.csv"),
            (ImportBatch.Kind.MATERIAL_ISSUE, "投料记录.csv"),
            (ImportBatch.Kind.WORK_REPORT, "报工记录.csv"),
        ]
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                for kind, filename in cases:
                    content = (template_dir / filename).read_bytes()
                    uploaded = SimpleUploadedFile(filename, content, "text/csv")
                    batch = create_import_batch(
                        period=self.period,
                        kind=kind,
                        uploaded_file=uploaded,
                        user=self.user,
                    )
                    process_batch(batch)

        self.assertEqual(self.period.work_orders.count(), 1)
        self.assertEqual(self.period.material_issues.count(), 1)
        self.assertEqual(self.period.work_reports.count(), 1)

        wo = self.period.work_orders.get()
        self.assertEqual(wo.order_no, "WO-202607-001")
        self.assertEqual(wo.line_code, "LN-01")
        self.assertEqual(wo.use_bom, True)
        self.assertTrue(wo.planned_start)

        issue = self.period.material_issues.get()
        self.assertEqual(issue.material_code, "RM-001")
        self.assertEqual(issue.issue_quantity, Decimal("500"))
        self.assertEqual(issue.returned_quantity, Decimal("50"))
        self.assertEqual(issue.actual_issue_quantity, Decimal("450"))

        report = self.period.work_reports.get()
        self.assertEqual(report.material_code, "FG-001")
        self.assertEqual(report.quantity, Decimal("1000"))
        self.assertEqual(report.work_hours, Decimal("8"))
        self.assertEqual(report.work_order_no, "WO-202607-001")

        # 三个台账页面可访问
        self.client.force_login(self.user)
        for url_name in (
            "work_order_imports",
            "material_issue_imports",
            "work_report_imports",
        ):
            response = self.client.get(
                reverse(url_name, kwargs={"period_id": self.period.id})
            )
            self.assertEqual(response.status_code, 200)

    def test_opening_inventory_import_feeds_first_period_valuation(self):
        period = AccountingPeriod.objects.create(month=date(2026, 1, 1))
        template_dir = Path(settings.BASE_DIR) / "sample_templates"
        content = (template_dir / "期初存货.csv").read_bytes()
        uploaded = SimpleUploadedFile("期初存货.csv", content, "text/csv")
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                batch = create_import_batch(
                    period=period,
                    kind=ImportBatch.Kind.OPENING_INVENTORY,
                    uploaded_file=uploaded,
                    user=self.user,
                )
                process_batch(batch)
        opening = OpeningInventoryBalance.objects.get(period=period)
        self.assertFalse(opening.is_system_generated)
        self.assertIsNone(opening.source_period)
        self.assertEqual(opening.quantity, Decimal("100"))
        self.assertEqual(opening.amount, Decimal("1000"))
        self.assertEqual(opening.unit_price, Decimal("10"))
        self.assertEqual(opening.opening_date, date(2026, 7, 31))
        self.assertEqual(opening.category_name, "原材料")
        self.assertEqual(opening.warehouse, "原料仓")
        self.assertEqual(opening.batch_no, "B20260701")
        self.assertEqual(opening.uom, "kg")

        run = run_settlement(period)
        self.assertTrue(run.converged)
        valuation = period.inventory_valuations.get(material_code="OP-RM")
        self.assertEqual(valuation.opening_quantity, Decimal("100"))
        self.assertEqual(valuation.opening_amount, Decimal("1000"))
        self.assertEqual(valuation.ending_quantity, Decimal("100"))
        self.assertEqual(valuation.ending_amount, Decimal("1000"))
        self.assertFalse(
            period.reconciliations.filter(
                status=ReconciliationResult.Status.FAIL
            ).exists()
        )

    def test_manual_opening_inventory_overrides_carried_balance(self):
        june = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        inventory_batch = self.create_batch(
            june, ImportBatch.Kind.INVENTORY, "override-june-inventory"
        )
        raw = Material.objects.create(
            code="OVERRIDE-RM", name="覆盖原料", uom="kg"
        )
        opening_txn = self.create_txn(
            june,
            inventory_batch,
            txn_date=date(2026, 6, 1),
            document_no="OV-OPEN",
            business_type="期初结存",
            direction=InventoryTransaction.Direction.IN,
            material=raw,
            quantity=Decimal("10"),
        )
        opening_txn.unit_price = Decimal("10")
        opening_txn.amount = Decimal("100")
        opening_txn.price_source = "原始金额"
        opening_txn.save(update_fields=["unit_price", "amount", "price_source"])
        june_run = run_settlement(june, self.user)
        self.assertTrue(june_run.converged)
        june.status = AccountingPeriod.Status.LOCKED
        june.save(update_fields=["status"])

        july = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        content = (
            "物料编码,物料名称,期初数量,期初单价,期初金额\n"
            "OVERRIDE-RM,覆盖原料,20,15,300\n"
        ).encode("utf-8-sig")
        uploaded = SimpleUploadedFile("期初存货.csv", content, "text/csv")
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                batch = create_import_batch(
                    period=july,
                    kind=ImportBatch.Kind.OPENING_INVENTORY,
                    uploaded_file=uploaded,
                    user=self.user,
                )
                process_batch(batch)
        july_run = run_settlement(july, self.user)
        self.assertTrue(july_run.converged)
        opening = OpeningInventoryBalance.objects.get(
            period=july, material_code="OVERRIDE-RM"
        )
        self.assertFalse(opening.is_system_generated)
        self.assertIsNone(opening.source_period)
        self.assertEqual(opening.quantity, Decimal("20"))
        self.assertEqual(opening.amount, Decimal("300"))
        valuation = july.inventory_valuations.get(material_code="OVERRIDE-RM")
        self.assertEqual(valuation.opening_quantity, Decimal("20"))
        self.assertEqual(valuation.opening_amount, Decimal("300"))
        self.assertEqual(valuation.ending_quantity, Decimal("20"))
        self.assertEqual(valuation.ending_amount, Decimal("300"))

    def test_ending_inventory_is_calculated_from_opening_and_transactions(self):
        period = AccountingPeriod.objects.create(month=date(2026, 1, 1))
        template_dir = Path(settings.BASE_DIR) / "sample_templates"
        content = (template_dir / "期初存货.csv").read_bytes()
        uploaded = SimpleUploadedFile("期初存货.csv", content, "text/csv")
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                batch = create_import_batch(
                    period=period,
                    kind=ImportBatch.Kind.OPENING_INVENTORY,
                    uploaded_file=uploaded,
                    user=self.user,
                )
                process_batch(batch)
        inventory_batch = self.create_batch(
            period, ImportBatch.Kind.INVENTORY, "ending-inventory-in"
        )
        self.create_txn(
            period,
            inventory_batch,
            txn_date=date(2026, 1, 15),
            document_no="GR-END",
            business_type="采购入库",
            direction=InventoryTransaction.Direction.IN,
            material=Material.objects.get(code="OP-RM"),
            quantity=Decimal("50"),
        )
        run = run_settlement(period)
        self.assertTrue(run.converged)
        # 系统自动生成期末行（无人工期末的物料按 期初+入库−出库 生成）
        sys_ending = EndingInventoryBalance.objects.filter(
            period=period, is_system_generated=True
        )
        self.assertTrue(sys_ending.exists())
        self.assertTrue(
            sys_ending.filter(material_code="OP-RM").exists()
        )
        valuation = period.inventory_valuations.get(material_code="OP-RM")
        self.assertEqual(valuation.ending_quantity, Decimal("150"))
        self.assertEqual(valuation.ending_amount, Decimal("1500"))
        self.assertFalse(period.exceptions.filter(category="期末存货").exists())

        # 人工上传期末存货可覆盖自动计算数（同物料）
        template_dir = Path(settings.BASE_DIR) / "sample_templates"
        content = (template_dir / "期末存货.csv").read_bytes()
        uploaded = SimpleUploadedFile("期末存货.csv", content, "text/csv")
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                batch = create_import_batch(
                    period=period,
                    kind=ImportBatch.Kind.ENDING_INVENTORY,
                    uploaded_file=uploaded,
                    user=self.user,
                )
                process_batch(batch)
        run = run_settlement(period)
        self.assertTrue(run.converged)
        ending = EndingInventoryBalance.objects.get(
            period=period, is_system_generated=False
        )
        self.assertEqual(ending.quantity, Decimal("100"))
        self.assertEqual(ending.amount, Decimal("1000"))
        self.assertEqual(ending.batch_no, "B20260701")

    def test_document_led_matching_workbench_filters_manual_and_ignore_modes(self):
        inventory_batch = self.create_batch(
            self.period, ImportBatch.Kind.INVENTORY, "workbench-inventory"
        )
        purchase_batch = self.create_batch(
            self.period,
            ImportBatch.Kind.PURCHASE_INVOICE,
            "workbench-purchase",
        )
        sales_batch = self.create_batch(
            self.period,
            ImportBatch.Kind.SALES_INVOICE,
            "workbench-sales",
        )
        order_batch = self.create_batch(
            self.period,
            ImportBatch.Kind.WORK_ORDER,
            "workbench-order",
        )
        raw = Material.objects.create(code="WB-RM", name="工作台原料", uom="kg")
        product = Material.objects.create(
            code="WB-FG", name="工作台产品", uom="件"
        )
        receipt = self.create_txn(
            self.period,
            inventory_batch,
            txn_date=date(2026, 8, 2),
            document_no="GR-WB",
            business_type="采购入库",
            direction=InventoryTransaction.Direction.IN,
            material=raw,
            quantity=Decimal("10"),
            supplier_code="SUP-WB",
        )
        sales_issue = self.create_txn(
            self.period,
            inventory_batch,
            txn_date=date(2026, 8, 20),
            document_no="SO-WB",
            business_type="销售出库",
            direction=InventoryTransaction.Direction.OUT,
            material=product,
            quantity=Decimal("5"),
            customer_code="CUS-WB",
            order_no="SALE-WB",
        )
        material_issue = self.create_txn(
            self.period,
            inventory_batch,
            txn_date=date(2026, 8, 10),
            document_no="ISS-WB",
            business_type="生产领料",
            direction=InventoryTransaction.Direction.OUT,
            material=raw,
            quantity=Decimal("3"),
            order_no="MO-WB",
            workshop="测试车间",
        )
        purchase_invoice = PurchaseInvoiceLine.objects.create(
            period=self.period,
            batch=purchase_batch,
            invoice_no="PI-WB",
            invoice_date=date(2026, 8, 5),
            invoice_period=date(2026, 8, 1),
            supplier_code="SUP-WB",
            material_code=raw.code,
            material_name=raw.name,
            quantity=Decimal("10"),
            unit_price_excl_tax=Decimal("8"),
            amount_excl_tax=Decimal("80"),
            remaining_quantity=Decimal("10"),
        )
        sales_invoice = SalesInvoiceLine.objects.create(
            period=self.period,
            batch=sales_batch,
            invoice_no="SI-WB",
            invoice_date=date(2026, 8, 21),
            invoice_period=date(2026, 8, 1),
            sales_order_no="SALE-WB",
            customer_code="CUS-WB",
            material_code=product.code,
            material_name=product.name,
            quantity=Decimal("5"),
            unit_price_excl_tax=Decimal("20"),
            amount_excl_tax=Decimal("100"),
            remaining_quantity=Decimal("5"),
        )
        order = WorkOrder.objects.create(
            period=self.period,
            batch=order_batch,
            order_no="MO-WB",
            production_dept_code="测试车间",
            plan_quantity=Decimal("5"),
        )
        auto_match_all(self.period)

        self.client.force_login(self.user)
        purchase_page = self.client.get(
            reverse("purchase_matches", kwargs={"period_id": self.period.id})
        )
        # 采购发票匹配页已按用户要求移除「发票号 / 明细ID」列（发票号是平台自动生成的 B 开头编号，
        # 用户不提供），故改为用仍展示的供应商 + 核销到的入库单号确认双边关系。
        self.assertContains(purchase_page, "SUP-WB")
        self.assertContains(purchase_page, "GR-WB")
        # 方案A：销售侧不再核销 → 核销工作台只剩采购侧；
        # 销售发票仍可上传，但不再核销库存流水（月末未开票按销售台账口径）
        sales_page = self.client.get(
            reverse("sales_uninvoiced", kwargs={"period_id": self.period.id})
        )
        self.assertEqual(sales_page.status_code, 200)
        inventory_page = self.client.get(
            reverse("inventory_matches", kwargs={"period_id": self.period.id}),
            {"document_no": "GR-WB"},
        )
        self.assertContains(inventory_page, "GR-WB")
        self.assertNotContains(inventory_page, "SO-WB")

        purchase_detail = reverse(
            "source_document_detail",
            kwargs={
                "period_id": self.period.id,
                "source_type": "purchase_invoice",
                "source_id": purchase_invoice.id,
            },
        )
        self.client.post(
            purchase_detail,
            {
                "action": "set_mode",
                "match_mode": MatchingMode.IGNORED,
                "ignore_reason": "测试忽略",
            },
        )
        purchase_invoice.refresh_from_db()
        self.assertEqual(purchase_invoice.match_mode, MatchingMode.IGNORED)
        self.assertFalse(
            MatchAllocation.objects.filter(
                purchase_invoice_line=purchase_invoice
            ).exists()
        )
        auto_match_all(self.period)
        self.assertFalse(
            MatchAllocation.objects.filter(
                purchase_invoice_line=purchase_invoice
            ).exists()
        )

        self.client.post(
            purchase_detail,
            {"action": "set_mode", "match_mode": MatchingMode.MANUAL},
        )
        self.client.post(
            purchase_detail,
            {
                "action": "add_match",
                "transaction_id": receipt.id,
                "matched_quantity": "10",
            },
        )
        self.assertTrue(
            MatchAllocation.objects.filter(
                purchase_invoice_line=purchase_invoice,
                inventory_transaction=receipt,
                is_manual=True,
            ).exists()
        )

        inventory_detail = reverse(
            "inventory_match_detail",
            kwargs={
                "period_id": self.period.id,
                "transaction_id": sales_issue.id,
            },
        )
        self.client.post(
            inventory_detail,
            {
                "action": "set_mode",
                "match_mode": MatchingMode.IGNORED,
                "ignore_reason": "本月忽略匹配",
            },
        )
        sales_issue.refresh_from_db()
        self.assertEqual(sales_issue.match_mode, MatchingMode.IGNORED)
        self.assertEqual(
            sales_issue.match_status,
            InventoryTransaction.MatchStatus.NOT_REQUIRED,
        )
        self.assertFalse(sales_issue.allocations.exists())
        self.assertEqual(sales_invoice.match_mode, MatchingMode.AUTO)

    def test_xlsx_inventory_upload(self):
        from openpyxl import Workbook

        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(
            ["业务日期", "单据号", "业务类型", "方向", "物料编码", "物料名称", "数量", "单位"]
        )
        worksheet.append(
            [date(2026, 8, 1), "GR-XLSX", "采购入库", "入库", "RM-X", "XLSX原料", 12, "kg"]
        )
        stream = BytesIO()
        workbook.save(stream)
        workbook.close()
        uploaded = SimpleUploadedFile(
            "inventory.xlsx",
            stream.getvalue(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                batch = create_import_batch(
                    period=self.period,
                    kind=ImportBatch.Kind.INVENTORY,
                    uploaded_file=uploaded,
                    user=self.user,
                )
                process_batch(batch)
        txn = self.period.inventory_transactions.get(document_no="GR-XLSX")
        self.assertEqual(txn.material_code, "RM-X")
        self.assertEqual(txn.quantity, Decimal("12"))

    def test_opening_detail_ledgers_can_only_be_uploaded_in_first_period(self):
        later = AccountingPeriod.objects.create(month=date(2026, 9, 1))
        content = (
            "原入库月份,原入库日期,原入库单号,物料编码,"
            "月初未到票数量,暂估单价\n"
            "2026-07,2026-07-20,OPEN-LATE,RM-LATE,10,2\n"
        ).encode("utf-8-sig")
        uploaded = SimpleUploadedFile(
            "opening-purchase.csv",
            content,
            "text/csv",
        )
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                with self.assertRaisesMessage(ValueError, "系统首个核算期间"):
                    create_import_batch(
                        period=later,
                        kind=ImportBatch.Kind.OPENING_PURCHASE_ACCRUAL,
                        uploaded_file=uploaded,
                        user=self.user,
                    )




class JiZhiWipQuantityTests(SettlementFixtureMixin, TestCase):
    """机织在产品重量法：本期投料台账没有该工单时按累计口径兜底（2026-09-29 修）。

    用户实测工单 32.29G-HF20250815G/HF20261259G：
      投料 14,607.65kg 只记在「投料进度表」/「工单明细」，投料记录台账（MaterialIssue）里 0 行；
      本期报工 2,487.3 米（本期完工 9.95%）、累计完工 13,137.6 米（52.55%）。
    修前：投料＝0 → 重量法在产品必然 0 → 766,624 元全进完工，完工单位成本 309.42 元/米。
    修后：在产品＝min(累计重量法, 计划 − 累计完工)＝25,000 − 13,137.6 ＝ 11,862.4 米。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="jz-wip-user", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.batch = self.create_batch(
            self.period, ImportBatch.Kind.WORK_ORDER, "jz-wip"
        )
        self.order = WorkOrder.objects.create(
            period=self.period,
            batch=self.batch,
            order_no="JZ-WIP-1",
            production_dept_code="JiZhi",
            plan_quantity=Decimal("25000"),
        )
        # BOM 单耗 41.34 ÷ 100 ＝ 0.4134 kg/米（机织米重）
        BomLine.objects.create(
            period=self.period,
            batch=self.batch,
            work_order_no="JZ-WIP-1",
            material_code="FG-JZ",
            material_name="机织面料",
            sub_item_code="Y-JZ-1",
            sub_item_name="预氧芳纶混纺纱线",
            qty_numerator=Decimal("41.34"),
            qty_denominator=Decimal("100"),
            uom="千克",
        )
        # 工单明细：累计完工（入库数量）13,137.6 米；累计投料 14,607.65 千克
        WorkOrderDetail.objects.create(
            period=self.period,
            batch=self.batch,
            order_no="JZ-WIP-1",
            plan_quantity=Decimal("25000"),
            issue_quantity=Decimal("14607.65"),
            receipt_quantity=Decimal("13137.6"),
            completion_status="超时未完成",
        )
        # 投料只记在投料进度表（MaterialIssue 里没有这张单）——原来就是这里漏掉
        IssueProgress.objects.create(
            period=self.period,
            batch=self.batch,
            work_order_no="JZ-WIP-1",
            material_code="Y-JZ-1",
            issue_quantity=Decimal("11963.65"),
            actual_issue_quantity=Decimal("14607.65"),
            uom="千克",
        )
        # 本期报工 2,487.3 米
        WorkReport.objects.create(
            period=self.period,
            batch=self.batch,
            report_time=datetime(2026, 7, 8, 8, 0),
            work_order_no="JZ-WIP-1",
            material_code="FG-JZ",
            material_name="机织面料",
            quantity=Decimal("2487.3"),
            uom="米",
        )
        # 期初在产品：数量 0、材料金额 666,224.33（与实测数据一致）
        OpeningWip.objects.create(
            period=self.period,
            batch=self.batch,
            order_no="JZ-WIP-1",
            quantity=Decimal("0"),
            material=Decimal("666224.33"),
        )

    def _run(self):
        from settlement.services.workorder_costing import calculate_work_order_costs

        return {
            r.work_order.order_no: r
            for r in calculate_work_order_costs(self.period)
        }

    def test_wip_uses_accumulated_issue_when_ledger_is_empty(self):
        result = self._run()["JZ-WIP-1"]
        # 在产品（用户 2026-09-30：机织在产品＝在产纱线重量 kg，不再折米）
        #  ＝ 米数法上限（计划 25,000 − 累计完工 13,137.6）＝11,862.4 米 × 米重 0.4134
        self.assertEqual(result.wip_quantity, Decimal("4903.916160")) 
        self.assertEqual(result.finished_quantity, Decimal("2487.3"))
        # 材料不再全进完工：完工只分到「完工耗料kg ÷（完工耗料kg＋在产品kg）」那部分
        self.assertGreater(result.wip_cost, result.finished_cost)
        self.assertGreater(result.wip_material, Decimal("0"))
        self.assertLess(result.finished_unit_cost, Decimal("100"))

    def test_no_wip_added_when_ledger_present(self):
        """本期投料台账有的工单仍走本期口径（行为不变）：投料被本期报工消耗完 → 在产品 0。"""
        MaterialIssue.objects.create(
            period=self.period,
            batch=self.batch,
            work_order_no="JZ-WIP-1",
            material_code="Y-JZ-1",
            material_name="预氧芳纶混纺纱线",
            material_category="纱线",
            issue_quantity=Decimal("1000"),
            returned_quantity=Decimal("0"),
            actual_issue_quantity=Decimal("1000"),
            issue_uom="千克",
        )
        result = self._run()["JZ-WIP-1"]
        # 期初 0 + 本期投料 1000 − 本期报工 2487.3×0.4134×1.075(=1105.5) ≤ 0 → 在产品 0
        self.assertEqual(result.wip_quantity, Decimal("0"))
        self.assertEqual(result.wip_cost, Decimal("0"))

    def test_wip_uses_opening_material_when_opening_quantity_missing(self):
        """期初在产品只有金额、没有数量时，按「金额 ÷ 投料单价」折算期初 kg（2026-09-30 修）。

        实测 32.29G-HF20260508：期初材料 533,156.37 元 ÷ 75.7798 = 7,035.60kg；
        本机 177 行期初在产品的「数量」列几乎全是空的，不折算的话重量法从 0kg 起步，
        在产品会被算成 0、成本全压到本期完工（该单实测：在产品 0、完工单位成本 99.26）。
        """
        from settlement.models import HistoricalPrice, MaterialIssue

        # 投料单价 100 元/kg（历史单价兜底即可，不必传 prices）
        HistoricalPrice.objects.create(
            period=self.period,
            material_code="Y-JZ-1",
            material_name="预氧芳纶混纺纱线",
            issue_price=Decimal("100"),
            receipt_price=Decimal("100"),
        )
        # 本月投料 4,731.75kg；报工 10,711 米（米重 0.432 → 消耗 4,974.19kg）
        MaterialIssue.objects.create(
            period=self.period,
            batch=self.batch,
            work_order_no="JZ-WIP-1",
            material_code="Y-JZ-1",
            material_name="预氧芳纶混纺纱线",
            material_category="纱线",
            issue_quantity=Decimal("4731.75"),
            returned_quantity=Decimal("0"),
            actual_issue_quantity=Decimal("4731.75"),
            issue_uom="千克",
        )
        WorkReport.objects.filter(work_order_no="JZ-WIP-1").update(
            quantity=Decimal("10711")
        )
        # 期初在产品：只有金额（7,000kg × 100 元），数量列是空的
        OpeningWip.objects.filter(order_no="JZ-WIP-1").update(
            quantity=Decimal("0"), material=Decimal("700000")
        )
        # 工单明细的入库数量是计划的 132.88%（实测同款数据）→ 不能当已完工用
        WorkOrderDetail.objects.filter(order_no="JZ-WIP-1").update(
            receipt_quantity=Decimal("33220.3")
        )
        result = self._run()["JZ-WIP-1"]
        # 重量法 (7,000 + 4,731.75 − 4,974.19) ＝ 6,757.56 kg，被「（计划 − 本月报工）× 米重」截到
        # (25,000 − 10,711) × 0.4134 ＝ 5,907.07 kg（用户 2026-09-30：在产品按在产纱线重量 kg）
        self.assertEqual(result.wip_quantity, Decimal("5907.072600"))
        self.assertEqual(result.finished_quantity, Decimal("10711"))
        self.assertGreater(result.wip_cost, result.finished_cost)
        self.assertLess(result.finished_unit_cost, Decimal("60"))

    def test_no_wip_for_cumulatively_finished_order(self):
        """累计完工 ≥ 计划（已完工的单）：兜底不会给它重新挂上在产品。"""
        WorkOrderDetail.objects.filter(order_no="JZ-WIP-1").update(
            receipt_quantity=Decimal("25000")
        )
        result = self._run()["JZ-WIP-1"]
        self.assertEqual(result.wip_quantity, Decimal("0"))

    def test_equivalent_quantity_uses_kg_for_jizhi(self):
        """折合产量口径（用户 2026-09-30）：机织（有 BOM 米重）以 kg 计，其它车间用各自数量。"""
        from settlement.services.workorder_costing import equivalent_quantity

        kg_map = {"JZ-WIP-1": Decimal("0.4134")}
        # 机织：完工 1,000 米 × 0.4134 ＝ 413.4 kg ＋ 0.5 × 在产品 400 kg ＝ 613.4
        self.assertEqual(
            equivalent_quantity("JiZhi", "JZ-WIP-1", Decimal("1000"), Decimal("400"), kg_map),
            Decimal("613.4"),
        )
        # 没米重的机织单退回原口径（米）
        self.assertEqual(
            equivalent_quantity("JiZhi", "JZ-NO-BOM", Decimal("1000"), Decimal("400"), kg_map),
            Decimal("1200"),
        )
        # 其它车间：完工 ＋ 0.5×在产品
        self.assertEqual(
            equivalent_quantity("RanZheng", "RZ-1", Decimal("1000"), Decimal("400"), kg_map),
            Decimal("1200"),
        )

    def test_backfill_material_applies_without_manual_adjustment(self):
        """BOM 反推 bug 回归（2026-09-30 修）：**没有人工调整**的工单也必须走 BOM 反推。

        原来 `reported` 过滤把系统报工数量传成 ZERO → 没有人工调整的工单全被滤掉
        （实测 311 张有报工的工单被滤成 0 张）→ 报了工没投料的单材料成本一直是 0。
        """
        from settlement.services.workorder_costing import calculate_work_order_costs

        # 一张"有报工、无投料、无期初"的染色单：材料成本应按 BOM 单耗×报工数量反推
        batch = self.create_batch(self.period, ImportBatch.Kind.WORK_ORDER, "bf-bug")
        WorkOrder.objects.create(
            period=self.period, batch=batch, order_no="BF-BUG-1",
            production_dept_code="RanZheng", plan_quantity=Decimal("1000"),
        )
        BomLine.objects.create(
            period=self.period, batch=batch, work_order_no="BF-BUG-1",
            material_code="P-BF", sub_item_code="Y-BF-1",
            qty_numerator=Decimal("2"), qty_denominator=Decimal("100"),
        )
        WorkReport.objects.create(
            period=self.period, batch=batch,
            report_time=datetime(2026, 7, 9, 8, 0),
            work_order_no="BF-BUG-1", quantity=Decimal("100"),
        )
        # 生产路径（run_settlement）会把库存结转的月末加权价传进来
        results = {
            r.work_order.order_no: r
            for r in calculate_work_order_costs(
                self.period, None, {"Y-BF-1": Decimal("50")}
            )
        }
        result = results.get("BF-BUG-1")
        self.assertIsNotNone(result)
        # 0.02 × 100 × 50 ＝ 100 元（修复前是 0）
        self.assertEqual(result.direct_material, Decimal("100.00"))
        self.assertEqual(result.finished_material, Decimal("100.00"))
        # 料是按报工量反推的 → **不确认在产品**（用户 2026-09-30：无投料-有报工怎么可能有在产品）
        self.assertEqual(result.wip_quantity, Decimal("0"))
        self.assertEqual(result.wip_cost, Decimal("0.00"))

    def test_material_code_trailing_zero_tolerance(self):
        """用户 2026-09-30（实测工单 18.441-1358059）：投料记录写 18.660、价源里是 18.66
        （同名同料）→ 也要取到价，否则整张单材料成本＝0。"""
        from settlement.services.workorder_costing import calculate_work_order_costs

        batch = self.create_batch(self.period, ImportBatch.Kind.WORK_ORDER, "code-fix")
        WorkOrder.objects.create(
            period=self.period, batch=batch, order_no="CODE-1",
            production_dept_code="RanZheng", plan_quantity=Decimal("200"),
        )
        MaterialIssue.objects.create(
            period=self.period, batch=batch, work_order_no="CODE-1",
            material_code="18.660", material_name="碳灰色间位芳纶缝纫线",
            material_category="缝纫线", issue_quantity=Decimal("125.43"),
            returned_quantity=Decimal("0"), actual_issue_quantity=Decimal("125.43"),
            issue_uom="千克",
        )
        WorkReport.objects.create(
            period=self.period, batch=batch,
            report_time=datetime(2026, 7, 9, 8, 0),
            work_order_no="CODE-1", quantity=Decimal("128.33"),
        )
        results = {
            r.work_order.order_no: r
            for r in calculate_work_order_costs(
                self.period, None, {"18.66": Decimal("108.20314119")}
            )
        }
        # 125.43 × 108.20314119 ＝ 13,571.92（精确匹配取不到价时按"去尾零"容错取到）
        self.assertEqual(results["CODE-1"].direct_material, Decimal("13571.92"))

    def test_case3_gap_backfill_with_threshold(self):
        """情形③ 差额反推（用户 2026-09-30）：无期初、无补投、报工>投料的单，缺多少补多少。

        缺口 ＝ BOM应有（单耗×报工×价） − 本期实领金额；阈值 100% ＝ 一律自动补料
        （用户要求"自动补料，不然低于标准成本太多"）。
        """
        from settlement.services.workorder_costing import (
            BOM_GAP_BACKFILL_MAX_PCT,
            calculate_work_order_costs,
        )

        self.assertEqual(BOM_GAP_BACKFILL_MAX_PCT, Decimal("100"))
        batch = self.create_batch(self.period, ImportBatch.Kind.WORK_ORDER, "gap-fix")
        prices = {"Y-GAP-1": Decimal("10")}
        # 单耗 0.02 kg/件、报工 100 件 → BOM应有 2kg ＝ 20 元
        for order_no, issue_qty, report_qty in (
            ("GAP-OK", Decimal("1.5"), Decimal("100")),   # 实领 15 元 → 缺口 5 元 → 补到 20
            ("GAP-BIG", Decimal("0.2"), Decimal("100")),  # 实领 2 元 → 缺口 90% → 也补到 20
        ):
            WorkOrder.objects.create(
                period=self.period, batch=batch, order_no=order_no,
                production_dept_code="RanZheng", plan_quantity=Decimal("200"),
            )
            BomLine.objects.create(
                period=self.period, batch=batch, work_order_no=order_no,
                material_code="P-GAP", sub_item_code="Y-GAP-1",
                qty_numerator=Decimal("2"), qty_denominator=Decimal("100"),
            )
            MaterialIssue.objects.create(
                period=self.period, batch=batch, work_order_no=order_no,
                material_code="Y-GAP-1", material_name="测试纱线",
                material_category="纱线", issue_quantity=issue_qty,
                returned_quantity=Decimal("0"), actual_issue_quantity=issue_qty,
                issue_uom="千克",
            )
            WorkReport.objects.create(
                period=self.period, batch=batch,
                report_time=datetime(2026, 7, 9, 8, 0),
                work_order_no=order_no, quantity=report_qty,
            )
        results = {
            r.work_order.order_no: r
            for r in calculate_work_order_costs(self.period, None, prices)
        }
        # 实领 1.5×10 ＝ 15，应有 0.02×100×10 ＝ 20 → 补 5 → 材料成本 20
        self.assertEqual(results["GAP-OK"].direct_material, Decimal("20.00"))
        # 缺口率 90% 也照补（用户 2026-09-30：自动补料，不然低于标准太多）→ 材料成本也是 20
        self.assertEqual(results["GAP-BIG"].direct_material, Decimal("20.00"))

        # 反推进来的料全部进完工（不按数量比例给在产品）
        self.assertEqual(results["GAP-OK"].wip_material, Decimal("0.00"))


class TeFangWipQuantityTests(SettlementFixtureMixin, TestCase):
    """纺纱在产品口径（用户 2026-09-30 实测 RAJ031A-HF20260603）。

    只有**真·粗纱工单（BC- 开头）**在产品＝0；**用粗纱的下游纺纱工单**照常算在产品：
    本月有投料台账 → 投料−完工；本月没有台账 → 累计口径（投料进度表/工单明细的累计投料 − 累计完工）。
    修前把"用粗纱的下游单"也清零，导致 13.4 万料全压进 32.8 的完工量、单位成本 4,993 元（标准 149）。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="tf-wip-user", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.batch = self.create_batch(
            self.period, ImportBatch.Kind.WORK_ORDER, "tf-wip"
        )

    def _mk_order(self, order_no, plan=Decimal("9280")):
        return WorkOrder.objects.create(
            period=self.period, batch=self.batch, order_no=order_no,
            production_dept_code="tefang", plan_quantity=plan,
        )

    def test_over_issued_order_keeps_wip_even_when_finish_rate_ok(self):
        """用户 2026-09-30：**超投料的单不清零**（视同完工只处理"尾料"，不处理"留给在产的料"）。

        实测 13.554-HF20260725：计划 400、报工 400.62（完工率 100% ≥70%）、投料 772.4 kg
        （应有只 400.62 kg）→ 应按纺纱口径保留在产品 372.4，而不是全压进完工。
        """
        from settlement.models import HistoricalPrice
        from settlement.services.workorder_costing import calculate_work_order_costs

        self._mk_order("TF-OVER-1", plan=Decimal("400"))
        HistoricalPrice.objects.create(
            period=self.period, material_code="R-A",
            issue_price=Decimal("100"), receipt_price=Decimal("100"),
        )
        BomLine.objects.create(
            period=self.period, batch=self.batch, work_order_no="TF-OVER-1",
            material_code="P-A", sub_item_code="R-A",
            qty_numerator=Decimal("1"), qty_denominator=Decimal("1"),
        )
        # 报工 400（完工率 100%）＋ 完成状态已完工 → 原口径会视同完工清零
        WorkReport.objects.create(
            period=self.period, batch=self.batch,
            report_time=datetime(2026, 7, 5, 9, 0),
            work_order_no="TF-OVER-1", quantity=Decimal("400"),
        )
        WorkOrderDetail.objects.create(
            period=self.period, batch=self.batch, order_no="TF-OVER-1",
            plan_quantity=Decimal("400"), issue_quantity=Decimal("772.4"),
            receipt_quantity=Decimal("400"), completion_status="超时完成",
        )
        # 投料台账 772.4（＝报工所需的 1.93 倍）
        MaterialIssue.objects.create(
            period=self.period, batch=self.batch, work_order_no="TF-OVER-1",
            material_code="R-A", material_name="料", material_category="纱线",
            issue_quantity=Decimal("772.4"), returned_quantity=Decimal("0"),
            actual_issue_quantity=Decimal("772.4"), issue_uom="千克",
        )
        results = {
            r.work_order.order_no: r
            for r in calculate_work_order_costs(self.period, None, {"R-A": Decimal("100")})
        }
        result = results["TF-OVER-1"]
        # 纺纱在产品 ＝ 投料 − 完工 ＝ 372.4（超投的部分留给在产）
        self.assertEqual(result.wip_quantity, Decimal("372.400000"))
        # 完工材料 ＝ 应有（报工 400 × 单耗 1 × 价 100）＝ 40,000，余料 37,240 留在产品
        self.assertEqual(result.finished_material, Decimal("40000.00"))
        self.assertEqual(result.wip_material, Decimal("37240.00"))

    def test_mixed_bom_order_keeps_wip_from_accumulated_issue(self):
        from settlement.services.workorder_costing import calculate_work_order_costs

        self._mk_order("TF-WIP-1")
        # BOM 是**混料**（粗纱＋其它料，与 RAJ031A-HF20260603 同款）→ 必须有在产品
        BomLine.objects.create(
            period=self.period, batch=self.batch, work_order_no="TF-WIP-1",
            material_code="P-TF", sub_item_code="BC-16.0420ST",
            qty_numerator=Decimal("24"), qty_denominator=Decimal("100"),
        )
        BomLine.objects.create(
            period=self.period, batch=self.batch, work_order_no="TF-WIP-1",
            material_code="P-TF2", sub_item_code="R11.101",
            qty_numerator=Decimal("10"), qty_denominator=Decimal("100"),
        )
        # 本月没有投料台账；累计：投料进度表 10,147.40、工单明细入库 7,495.55
        IssueProgress.objects.create(
            period=self.period, batch=self.batch, work_order_no="TF-WIP-1",
            material_code="BC-16.0420ST", issue_quantity=Decimal("10147.40"),
            actual_issue_quantity=Decimal("10147.40"), uom="千克",
        )
        WorkOrderDetail.objects.create(
            period=self.period, batch=self.batch, order_no="TF-WIP-1",
            plan_quantity=Decimal("9280"), issue_quantity=Decimal("10147.40"),
            receipt_quantity=Decimal("7495.55"), completion_status="超时完成",
        )
        WorkReport.objects.create(
            period=self.period, batch=self.batch,
            report_time=datetime(2026, 7, 1, 9, 0),
            work_order_no="TF-WIP-1", quantity=Decimal("32.80"),
        )
        # 期初在产品（料是上期投的）：材料 133,713.26 —— 与实测同款
        OpeningWip.objects.create(
            period=self.period, batch=self.batch, order_no="TF-WIP-1",
            quantity=Decimal("0"), material=Decimal("133713.26"),
            wages=Decimal("5124.37"),
        )
        results = {
            r.work_order.order_no: r
            for r in calculate_work_order_costs(self.period)
        }
        result = results["TF-WIP-1"]
        # 在产品 ＝ 累计投料 10,147.40 − 累计完工 7,495.55 ＝ 2,651.85（修前是 0）
        self.assertEqual(result.wip_quantity, Decimal("2651.85"))
        self.assertEqual(result.finished_quantity, Decimal("32.80"))
        # 完工率 0.35% → 不触发"视同完工"清零；13.4 万料绝大部分留在在产品（修前全压进完工）
        self.assertGreater(result.wip_material, result.finished_material)
        # 材料＝期初结转 133,713.26（本测试没造投料/粗纱补投）
        self.assertEqual(result.direct_material, Decimal("133713.26"))

    def test_finished_material_backfilled_from_bom(self):
        """用户 2026-09-30：上月有在产、本月没投料但有报工的单 —— 完工材料按
        「完工数量 × BOM 单耗 × 价」反推，余额留在在产品（不再按数量比例摊一点）。
        """
        from settlement.services.workorder_costing import calculate_work_order_costs

        self._mk_order("TF-BOM-1", plan=Decimal("1000"))
        # BOM 单耗 0.5、单价 100 → 报工 100 的应有料 ＝ 50 × 100 ＝ 5,000
        BomLine.objects.create(
            period=self.period, batch=self.batch, work_order_no="TF-BOM-1",
            material_code="P-A", sub_item_code="R-A",
            qty_numerator=Decimal("5"), qty_denominator=Decimal("10"),
        )
        IssueProgress.objects.create(
            period=self.period, batch=self.batch, work_order_no="TF-BOM-1",
            material_code="R-A", issue_quantity=Decimal("800"),
            actual_issue_quantity=Decimal("800"), uom="千克",
        )
        WorkOrderDetail.objects.create(
            period=self.period, batch=self.batch, order_no="TF-BOM-1",
            plan_quantity=Decimal("1000"), issue_quantity=Decimal("800"),
            receipt_quantity=Decimal("200"), completion_status="超时未完成",
        )
        WorkReport.objects.create(
            period=self.period, batch=self.batch,
            report_time=datetime(2026, 7, 2, 9, 0),
            work_order_no="TF-BOM-1", quantity=Decimal("100"),
        )
        OpeningWip.objects.create(
            period=self.period, batch=self.batch, order_no="TF-BOM-1",
            quantity=Decimal("0"), material=Decimal("80000"),
        )
        results = {
            r.work_order.order_no: r
            for r in calculate_work_order_costs(
                self.period, None, {"R-A": Decimal("100")}
            )
        }
        result = results["TF-BOM-1"]
        # 在产品 ＝ 累计投料 800 − 入库 200 ＝ 600
        self.assertEqual(result.wip_quantity, Decimal("600"))
        # 完工材料 ＝ 100 × 0.5 × 100 ＝ 5,000（旧口径是 80,000×100/700 ≈ 11,428.57）
        self.assertEqual(result.finished_material, Decimal("5000.00"))
        self.assertEqual(result.wip_material, Decimal("75000.00"))

    def test_real_roving_order_still_has_no_wip(self):
        """真·粗纱工单（BC- 开头）本来就不报工、不挂在产品。"""
        from settlement.services.workorder_costing import calculate_work_order_costs

        self._mk_order("BC-16.0420ST-9001", plan=Decimal("100"))
        WorkReport.objects.create(
            period=self.period, batch=self.batch,
            report_time=datetime(2026, 7, 1, 9, 0),
            work_order_no="BC-16.0420ST-9001", quantity=Decimal("100"),
        )
        results = {
            r.work_order.order_no: r
            for r in calculate_work_order_costs(self.period)
        }
        self.assertEqual(results["BC-16.0420ST-9001"].wip_quantity, Decimal("0"))

    def test_all_roving_bom_order_has_no_wip(self):
        """用户 2026-09-30：**BOM 100% 只有粗纱**的工单没有在产品（混料的才有）。"""
        from settlement.services.workorder_costing import calculate_work_order_costs

        self._mk_order("TF-ALLROV-1", plan=Decimal("1000"))
        BomLine.objects.create(
            period=self.period, batch=self.batch, work_order_no="TF-ALLROV-1",
            material_code="P-TF", sub_item_code="BC-16.0026CS",
            qty_numerator=Decimal("1"), qty_denominator=Decimal("1"),
        )
        IssueProgress.objects.create(
            period=self.period, batch=self.batch, work_order_no="TF-ALLROV-1",
            material_code="BC-16.0026CS", issue_quantity=Decimal("800"),
            actual_issue_quantity=Decimal("800"), uom="千克",
        )
        WorkOrderDetail.objects.create(
            period=self.period, batch=self.batch, order_no="TF-ALLROV-1",
            plan_quantity=Decimal("1000"), issue_quantity=Decimal("800"),
            receipt_quantity=Decimal("200"), completion_status="超时未完成",
        )
        WorkReport.objects.create(
            period=self.period, batch=self.batch,
            report_time=datetime(2026, 7, 1, 9, 0),
            work_order_no="TF-ALLROV-1", quantity=Decimal("200"),
        )
        results = {
            r.work_order.order_no: r
            for r in calculate_work_order_costs(self.period)
        }
        self.assertEqual(results["TF-ALLROV-1"].wip_quantity, Decimal("0"))


class PulpLedgerImportTests(TestCase):
    """浆粕流水（原料+成品通用模板）：按业务类型分流导入。"""

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="ledger-user", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))

    def test_ledger_splits_by_business_type(self):
        content = (
            "创建时间,物料编号,物料名称,分类名称,物料规格,仓库名称,批次号,操作数量,单位名称,单价,金额,业务类型,车间,备注\n"
            "2026-07-31,R82.101,对位芳纶回收料-乱丝,原材料,,原料仓,,1000,千克,20,20000,领料出库,浆粕车间,用于芳纶浆粕/微纤\n"
            "2026-07-31,R82.101,对位芳纶回收料-乱丝,原材料,,原料仓,,500,千克,,,领料出库,浆粕车间,用于芳纶浆粕/微纤\n"
            "2026-07-31,21.03,21.03 芳纶浆粕 DFP-103,产成品,,成品仓,,2000,千克,,,生产入库,浆粕车间,7月生产入库\n"
            "2026-07-31,21.03,21.03 芳纶浆粕 DFP-103,产成品,,成品仓,,300,千克,,,销售出库,浆粕车间,7月销售出库\n"
            "2026-07-31,示例-R99,示例物料,原材料,,,,1,千克,,,领料出库,,示例行\n"
        ).encode("utf-8-sig")
        uploaded = SimpleUploadedFile("pulp_ledger.csv", content, "text/csv")
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                batch = create_import_batch(
                    period=self.period,
                    kind=ImportBatch.Kind.PULP_LEDGER,
                    uploaded_file=uploaded,
                    user=self.user,
                )
                process_batch(batch)
        self.assertEqual(batch.row_count, 4)
        # 领料出库 2 行 → PulpMaterial；生产入库/销售出库 2 行 → InventoryTransaction
        self.assertEqual(self.period.pulp_materials.count(), 2)
        self.assertEqual(
            self.period.inventory_transactions.filter(
                business_type__in=["生产入库", "销售出库"]
            ).count(),
            2,
        )
        HistoricalPrice.objects.create(
            period=self.period, material_code="R82.101",
            material_name="对位芳纶回收料-乱丝",
            issue_price=Decimal("20"), receipt_price=Decimal("18"),
        )

        # 材料成本只统计领料出库行：金额用用户值 + 无金额行（金额=0，结算时按历史单价）
        costs = collect_pulp_costs(self.period)
        self.assertEqual(costs["material_fanglun"], Decimal("30000"))

    def test_ledger_blank_business_type_skipped(self):
        # 业务类型为空的行（空行/说明行）自动跳过，不算未导入
        content = (
            "创建时间,物料编号,物料名称,操作数量,业务类型,备注\n"
            "2026-07-31,R82.101,对位芳纶回收料-乱丝,1000,领料出库,用于芳纶浆粕/微纤\n"
            "2026-07-31,21.03,21.03 芳纶浆粕 DFP-103,2000,生产入库,\n"
            ",,,,,合计\n"
        ).encode("utf-8-sig")
        uploaded = SimpleUploadedFile("pulp_ledger.csv", content, "text/csv")
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                batch = create_import_batch(
                    period=self.period,
                    kind=ImportBatch.Kind.PULP_LEDGER,
                    uploaded_file=uploaded,
                    user=self.user,
                )
                process_batch(batch)
        self.assertEqual(batch.row_count, 2)
        self.assertEqual(batch.error_count, 0)
        self.assertEqual(self.period.pulp_materials.count(), 1)

    def test_ledger_unsupported_business_type_raises(self):
        content = (
            "创建时间,物料编号,物料名称,操作数量,业务类型,备注\n"
            "2026-07-31,R82.101,对位芳纶回收料-乱丝,1000,采购入库,\n"
        ).encode("utf-8-sig")
        uploaded = SimpleUploadedFile("pulp_ledger.csv", content, "text/csv")
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                batch = create_import_batch(
                    period=self.period,
                    kind=ImportBatch.Kind.PULP_LEDGER,
                    uploaded_file=uploaded,
                    user=self.user,
                )
                with self.assertRaises(ValueError):
                    process_batch(batch)

class PulpMaterialImportTests(TestCase):
    """浆粕原料流水：上传即材料成本来源；未上传时浆粕车间不能结转。"""

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="pulp-user", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))

    def _import_csv(self, content):
        uploaded = SimpleUploadedFile("pulp_material.csv", content, "text/csv")
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                batch = create_import_batch(
                    period=self.period,
                    kind=ImportBatch.Kind.PULP_MATERIAL,
                    uploaded_file=uploaded,
                    user=self.user,
                )
                process_batch(batch)
        return batch

    def test_import_and_material_allocation(self):
        # 用户格式：创建时间/物料编号/物料名称/操作数量/业务类型/（可选）金额
        content = (
            "创建时间,物料编号,物料名称,操作数量,业务类型,金额,备注\n"
            "2026-07-31,R82.101,对位芳纶回收料-乱丝,1000,领料出库,,浆粕车间用于芳纶浆粕/微纤\n"
            "2026-07-31,R82.101,对位芳纶回收料-乱丝,500,领料出库,15000,浆粕车间用于芳纶浆粕/微纤\n"
"2026-07-31,R84.101,国产聚丙烯腈丝束,2000,领料出库,,用于腈纶浆粕/微纤\n"
            "2026-07-31,R84.101,国产聚丙烯腈丝束,3000,生产入库,,浆粕车间成品\n"
            "2026-07-31,示例-R99,示例物料,1,领料出库,,示例行\n"
        ).encode("utf-8-sig")
        batch = self._import_csv(content)
        self.assertEqual(batch.row_count, 4)
        mats = list(self.period.pulp_materials.all())
        self.assertEqual(len(mats), 4)
        # 用户金额列优先；无金额行 amount=0，结算时按历史单价定价
        self.assertEqual(mats[1].amount, Decimal("15000"))
        self.assertEqual(mats[0].amount, Decimal("0"))
        self.assertEqual(mats[3].business_type, "生产入库")

        # 历史单价：R82.101 出库价 20（优先）；R84.101 只有入库价 10.5（兜底）
        HistoricalPrice.objects.create(
            period=self.period, material_code="R82.101",
            material_name="对位芳纶回收料-乱丝",
            issue_price=Decimal("20"), receipt_price=Decimal("18"),
        )
        HistoricalPrice.objects.create(
            period=self.period, material_code="R84.101",
            material_name="国产聚丙烯腈丝束",
            issue_price=None, receipt_price=Decimal("10.5"),
        )

        WorkshopCostPool.objects.create(
            period=self.period, batch=batch, workshop="浆粕车间",
            water_power=1000, wages=2000,
        )
        for code, name, qty in [
            ("21.03", "21.03 芳纶浆粕 DFP-103", 2000),
            ("23.06", "23.06 腈纶浆粕 JLP-505", 3000),
        ]:
            InventoryTransaction.objects.create(
                period=self.period, batch=batch, material_code=code,
                material_name=name, business_type="生产入库",
                txn_date=date(2026, 7, 31),
                direction=InventoryTransaction.Direction.IN, quantity=qty,
                unit_price=0, amount=0,
            )
        costs = collect_pulp_costs(self.period)
        # 芳纶：1000×20（系统出库价）+ 500 行用户金额 15000 = 35000
        self.assertEqual(costs["material_fanglun"], Decimal("35000"))
        # 腈纶：2000×10.5（入库价兜底）；生产入库 3000kg 不计入材料
        self.assertEqual(costs["material_jinglun"], Decimal("21000"))
        # 芳纶成品 21.03 得芳纶材料全额；腈纶成品 23.06 得腈纶材料全额
        self.assertEqual(costs["products"]["21.03"]["material"], Decimal("35000"))
        self.assertEqual(costs["products"]["23.06"]["material"], Decimal("21000"))
        # 费用按产量分摊：21.03 占 2000/5000，工资同理
        self.assertEqual(costs["products"]["21.03"]["fee"], Decimal("400"))
        self.assertEqual(costs["products"]["23.06"]["fee"], Decimal("600"))
        self.assertEqual(costs["products"]["21.03"]["wages"], Decimal("800"))
        self.assertEqual(costs["products"]["23.06"]["wages"], Decimal("1200"))

    def test_missing_upload_blocks_settlement(self):
        with self.assertRaises(ValueError):
            collect_pulp_costs(self.period)

    def test_unclassifiable_row_raises(self):
        content = (
            "创建时间,物料编号,物料名称,操作数量,业务类型,备注\n"
            "2026-07-31,R99.999,未知物料,100,领料出库,\n"
        ).encode("utf-8-sig")
        self._import_csv(content)
        with self.assertRaises(ValueError):
            collect_pulp_costs(self.period)

    def test_remark_exact_product_allocation(self):
        # 备注指明具体成品 → 材料全部分给该成品（不按产量比例摊）
        content = (
            "创建时间,物料编号,物料名称,操作数量,业务类型,金额,备注\n"
            "2026-07-31,R82.101,对位芳纶回收料-乱丝,1000,领料出库,20000,用于21.03芳纶浆粕\n"
            "2026-07-31,R11.519,对位芳纶长丝,500,领料出库,5000,用于22.01芳纶微纤\n"
        ).encode("utf-8-sig")
        batch = self._import_csv(content)
        self.assertEqual(batch.row_count, 2)
        for code, name, qty in [
            ("21.03", "21.03 芳纶浆粕 DFP-103", 2000),
            ("22.01", "22.01 芳纶微纤 AMF-100", 1000),
        ]:
            InventoryTransaction.objects.create(
                period=self.period, batch=batch, material_code=code,
                material_name=name, business_type="生产入库",
                txn_date=date(2026, 7, 31),
                direction=InventoryTransaction.Direction.IN, quantity=qty,
                unit_price=0, amount=0,
            )
        costs = collect_pulp_costs(self.period)
        self.assertEqual(costs["products"]["21.03"]["material"], Decimal("20000"))
        self.assertEqual(costs["products"]["22.01"]["material"], Decimal("5000"))

    def test_remark_category_splits_by_product_name(self):
        # "用于芳纶浆粕/微纤" → 名称含芳纶浆粕/芳纶微纤的成品（21.03/22.01）按产量比例摊；腈纶同理
        content = (
            "创建时间,物料编号,物料名称,操作数量,业务类型,金额,备注\n"
            "2026-07-31,R82.101,对位芳纶回收料-乱丝,3000,领料出库,30000,用于芳纶浆粕/微纤\n"
            "2026-07-31,R84.101,国产聚丙烯腈丝束,3000,领料出库,21000,用于腈纶浆粕/微纤\n"
        ).encode("utf-8-sig")
        batch = self._import_csv(content)
        self.assertEqual(batch.row_count, 2)
        for code, name, qty in [
            ("21.03", "21.03 芳纶浆粕 DFP-103", 2000),
            ("22.01", "22.01 芳纶微纤 AMF-100", 1000),
            ("23.06", "23.06 腈纶浆粕 JLP-505", 3000),
            ("24.01", "24.01 腈纶微纤 JMF-025", 1000),
        ]:
            InventoryTransaction.objects.create(
                period=self.period, batch=batch, material_code=code,
                material_name=name, business_type="生产入库",
                txn_date=date(2026, 7, 31),
                direction=InventoryTransaction.Direction.IN, quantity=qty,
                unit_price=0, amount=0,
            )
        costs = collect_pulp_costs(self.period)
        # 芳纶 30000 按产量 2000:1000 摊 → 21.03=20000, 22.01=10000
        self.assertEqual(costs["products"]["21.03"]["material"], Decimal("20000"))
        self.assertAlmostEqual(float(costs["products"]["22.01"]["material"]), 10000.0, places=2)
        # 腈纶 21000 按产量 3000:1000 摊 → 23.06=15750, 24.01=5250
        self.assertEqual(costs["products"]["23.06"]["material"], Decimal("15750"))
        self.assertEqual(costs["products"]["24.01"]["material"], Decimal("5250"))

    def test_remark_unrecognized_falls_back_by_material(self):
        # 备注无法识别（如"浆粕车间7月领用"）→ 按物料名称类别兜底：含"芳纶"→ 芳纶类成品按产量比例摊
        content = (
            "创建时间,物料编号,物料名称,操作数量,业务类型,金额,备注\n"
            "2026-07-31,R82.101,对位芳纶回收料-乱丝,1000,领料出库,20000,浆粕车间7月领用\n"
            "2026-07-31,R84.101,国产聚丙烯腈丝束,500,领料出库,10500,浆粕车间7月领用\n"
        ).encode("utf-8-sig")
        batch = self._import_csv(content)
        self.assertEqual(batch.row_count, 2)
        for code, name, qty in [
            ("21.03", "21.03 芳纶浆粕 DFP-103", 2000),
            ("23.06", "23.06 腈纶浆粕 JLP-505", 3000),
        ]:
            InventoryTransaction.objects.create(
                period=self.period, batch=batch, material_code=code,
                material_name=name, business_type="生产入库",
                txn_date=date(2026, 7, 31),
                direction=InventoryTransaction.Direction.IN, quantity=qty,
                unit_price=0, amount=0,
            )
        # 不报错：芳纶材料兜底给芳纶类 21.03，腈纶材料兜底给腈纶类 23.06
        costs = collect_pulp_costs(self.period)
        self.assertEqual(costs["products"]["21.03"]["material"], Decimal("20000"))
        self.assertEqual(costs["products"]["23.06"]["material"], Decimal("10500"))

    def test_no_price_raises(self):
        # 领料出库行既无金额也无历史单价 → 报错提示
        content = (
            "创建时间,物料编号,物料名称,操作数量,业务类型,备注\n"
            "2026-07-31,R82.101,对位芳纶回收料-乱丝,1000,领料出库,用于芳纶浆粕/微纤\n"
        ).encode("utf-8-sig")
        self._import_csv(content)
        with self.assertRaises(ValueError):
            collect_pulp_costs(self.period)



class ProfitStatementTests(SettlementFixtureMixin, TestCase):
    """损益表：台账逐行 + 费用按【出库月份】分别录入并按该月收入占比分摊 + 月×渠道×大类汇总。

    口径（用户 2026-09-14 确认）：
    - 收入只取「终端销售收入」；射阳兰邦销售收入只作列展示；
    - 销售毛利 = 终端销售收入 − 销售成本；
    - 7 项费用（税金及附加/管理费用/销售费用/财务费用/营业外收入/营业外支出/所得税费用）
      按出库月份分别录入，每行按「本行终端收入 ÷ 该行所属出库月份终端收入合计」分摊，
      尾差计入该月收入最大行 ⇒ 每月各项分摊合计 = 该月录入金额；
    - 营业利润 = 毛利 − 税及附加 − 三费；利润总额 = 营业利润 + 营业外收入 − 营业外支出；
      净利润 = 利润总额 − 所得税费用；
    - 汇总层级：月份 → 销售渠道（国内销售/国外销售/未分类）→ 销售大类，
      月份可按出库期间或开票期间切换。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="pl-tester", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.batch = self.create_batch(
            self.period, ImportBatch.Kind.SALES_LEDGER, "pl"
        )
        self.june = date(2026, 6, 1)

    _UNSET = object()

    def _line(
        self,
        row_no,
        *,
        revenue,
        cost=None,
        quantity=Decimal("10"),
        lb=None,
        out_month=None,
        invoice_month=_UNSET,
        channel="国内销售",
        category="纱线",
        customer="测试客户",
    ):
        return SalesLedgerLine.objects.create(
            period=self.period,
            batch=self.batch,
            row_no=row_no,
            out_month=out_month or self.period.month,
            invoice_status=SalesLedgerLine.InvoiceStatus.INVOICED,
            invoice_month=(
                self.period.month if invoice_month is self._UNSET else invoice_month
            ),
            channel=channel,
            company="江苏盐城兰邦特种纺织集团有限责任公司",
            order_no="SO2026070001",
            customer_code="C001",
            customer_name=customer,
            sales_category=category,
            material_code="21.03",
            material_name="21.03 芳纶浆粕 DFP-103",
            specification="DFP-103",
            uom="kg",
            terminal_amount=revenue,
            terminal_unit_price=Decimal("100") if revenue else None,
            lb_amount=lb,
            quantity=quantity,
            cost_amount=cost,
        )

    def _expense(self, month, **kwargs):
        return ProfitStatementExpense.objects.create(
            period=self.period, month=month, **kwargs
        )

    # —————————————— 分摊 ——————————————

    def test_expenses_allocated_within_own_month(self):
        # 6 月两行、7 月两行；两个月的费用各自分摊，互不串月
        self._line(1, revenue=Decimal("100"), cost=Decimal("10"), out_month=self.june)
        self._line(2, revenue=Decimal("300"), cost=Decimal("30"), out_month=self.june)
        self._line(3, revenue=Decimal("200"), cost=Decimal("20"))
        self._line(4, revenue=Decimal("600"), cost=Decimal("60"))
        self._expense(self.june, tax_surcharge=Decimal("40"), admin_expense=Decimal("80"))
        self._expense(self.period.month, tax_surcharge=Decimal("800"),
                      admin_expense=Decimal("1600"))

        rows = build_profit_statement(self.period)["rows"]
        self.assertEqual(
            [row["tax_surcharge"] for row in rows],
            [Decimal("10.00"), Decimal("30.00"), Decimal("200.00"), Decimal("600.00")],
        )
        self.assertEqual(
            [row["admin_expense"] for row in rows],
            [Decimal("20.00"), Decimal("60.00"), Decimal("400.00"), Decimal("1200.00")],
        )
        # 各月分摊合计 = 该月录入金额
        self.assertEqual(
            sum(row["tax_surcharge"] for row in rows if row["out_month"] == "2026-06"),
            Decimal("40.00"),
        )
        self.assertEqual(
            sum(row["tax_surcharge"] for row in rows if row["out_month"] == "2026-07"),
            Decimal("800.00"),
        )

    def test_month_without_entry_gets_no_allocation(self):
        # 没录费用的月份：分摊全为 0，营业利润 = 毛利
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"), out_month=self.june)
        self._line(2, revenue=Decimal("1000"), cost=Decimal("600"))
        self._expense(self.period.month, admin_expense=Decimal("100"))

        rows = build_profit_statement(self.period)["rows"]
        self.assertEqual(rows[0]["admin_expense"], Decimal("0.00"))
        self.assertEqual(rows[0]["operating_profit"], Decimal("400.00"))
        self.assertEqual(rows[1]["admin_expense"], Decimal("100.00"))
        self.assertEqual(rows[1]["operating_profit"], Decimal("300.00"))

    def test_each_allocation_column_totals_to_entered_amount(self):
        # 三行收入不能被整除 → 每列分摊合计仍必须等于录入金额（尾差计入收入最大行）
        for row_no in (1, 2, 3):
            self._line(row_no, revenue=Decimal("333.33"), cost=Decimal("100"))
        expense = self._expense(
            self.period.month,
            tax_surcharge=Decimal("100"),
            admin_expense=Decimal("200"),
            selling_expense=Decimal("300"),
            finance_expense=Decimal("400"),
            non_operating_income=Decimal("50"),
            non_operating_expense=Decimal("60"),
            income_tax=Decimal("70"),
        )

        data = build_profit_statement(self.period)
        for field, _label in ALLOCATION_FIELDS:
            entered = getattr(expense, field)
            self.assertEqual(
                sum(row[field] for row in data["rows"]),
                entered,
                msg=f"{field} 分摊合计应等于录入金额",
            )
            self.assertEqual(data["totals"][field], entered)

        tax_values = [row["tax_surcharge"] for row in data["rows"]]
        self.assertEqual(sum(tax_values), Decimal("100.00"))
        self.assertEqual(tax_values.count(Decimal("33.34")), 1)

    def test_gross_to_net_formulas(self):
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"))
        self._expense(
            self.period.month,
            tax_surcharge=Decimal("10"),
            admin_expense=Decimal("20"),
            selling_expense=Decimal("30"),
            finance_expense=Decimal("40"),
            non_operating_income=Decimal("5"),
            non_operating_expense=Decimal("7"),
            income_tax=Decimal("8"),
        )

        data = build_profit_statement(self.period)
        row = data["rows"][0]
        self.assertEqual(row["gross"], Decimal("400.00"))
        self.assertEqual(row["gross_margin"], Decimal("40.00"))
        self.assertEqual(row["operating_profit"], Decimal("300.00"))
        self.assertEqual(row["total_profit"], Decimal("298.00"))
        self.assertEqual(row["net_profit"], Decimal("290.00"))
        self.assertEqual(row["net_margin"], Decimal("29.00"))

        totals = data["totals"]
        self.assertEqual(totals["revenue"], Decimal("1000.00"))
        self.assertEqual(totals["cost"], Decimal("600.00"))
        self.assertEqual(totals["gross"], Decimal("400.00"))
        self.assertEqual(totals["operating_profit"], Decimal("300.00"))
        self.assertEqual(totals["total_profit"], Decimal("298.00"))
        self.assertEqual(totals["net_profit"], Decimal("290.00"))
        self.assertEqual(totals["gross_margin"], Decimal("40.00"))
        self.assertEqual(totals["net_margin"], Decimal("29.00"))

    def test_zero_terminal_revenue_rows_are_not_allocated(self):
        # 只填了射阳兰邦收入的行：收入为 0，不分摊费用，毛利率/净利润率显示为空
        self._line(1, revenue=Decimal("0"), cost=None, lb=Decimal("500"))
        self._expense(self.period.month, admin_expense=Decimal("100"))

        data = build_profit_statement(self.period)
        self.assertEqual(data["lb_only_rows"], 1)
        self.assertEqual(data["revenue_total"], Decimal("0"))
        row = data["rows"][0]
        self.assertEqual(row["admin_expense"], Decimal("0.00"))
        self.assertIsNone(row["gross_margin"])
        self.assertIsNone(row["net_margin"])

    # —————————————— 汇总（月 → 渠道 → 大类）——————————————

    def _summary(self, basis):
        data = build_profit_statement(self.period)
        return build_summary(data["rows"], basis)

    def test_summary_by_month_channel_and_category(self):
        # 7 月：国内-纱线 1000、国内-面料 500、国外-纱线 2000；6 月：国外-面料 100
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"), category="纱线")
        self._line(2, revenue=Decimal("500"), cost=Decimal("250"), category="面料")
        self._line(3, revenue=Decimal("2000"), cost=Decimal("1000"),
                   category="纱线", channel="国外销售")
        self._line(4, revenue=Decimal("100"), cost=Decimal("60"),
                   category="面料", channel="国外销售", out_month=self.june)

        summary = self._summary("out")
        by_key = {(r["month"], r["channel"], r["category"]): r for r in summary["rows"]}

        domestic_yarn = by_key[("2026-07", "国内销售", "纱线")]
        self.assertEqual(domestic_yarn["revenue"], Decimal("1000.00"))
        self.assertEqual(domestic_yarn["gross"], Decimal("400.00"))

        # 国内渠道小计 = 纱线 + 面料
        domestic_total = by_key[("2026-07", "国内销售", "小计")]
        self.assertEqual(domestic_total["revenue"], Decimal("1500.00"))
        self.assertEqual(domestic_total["cost"], Decimal("850.00"))
        self.assertEqual(domestic_total["gross"], Decimal("650.00"))

        # 7 月月度小计
        july_total = by_key[("2026-07", "小计", "")]
        self.assertEqual(july_total["revenue"], Decimal("3500.00"))

        # 总计
        self.assertEqual(summary["grand_total"]["revenue"], Decimal("3600.00"))
        self.assertEqual(summary["grand_total"]["cost"], Decimal("1910.00"))

        # 渠道顺序：国内销售 → 国外销售；大类按收入倒序
        channel_order = [r["channel"] for r in summary["rows"] if r["month"] == "2026-07"]
        self.assertEqual(channel_order[0], "国内销售")
        self.assertIn("国外销售", channel_order)
        categories = [r["category"] for r in summary["rows"]
                      if r["month"] == "2026-07" and r["channel"] == "国内销售"]
        self.assertEqual(categories, ["纱线", "面料", "小计"])

    def test_summary_puts_blank_channel_into_unclassified(self):
        self._line(1, revenue=Decimal("100"), cost=Decimal("50"), channel="")
        summary = self._summary("out")
        self.assertIn(("2026-07", "未分类", "纱线"),
                      {(r["month"], r["channel"], r["category"]) for r in summary["rows"]})

    def test_summary_basis_invoice_month(self):
        # 6 月发货、7 月开票 → 出库期间视图算 6 月；开票期间视图算 7 月；未开票归入「(未开票)」
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"),
                   out_month=self.june, invoice_month=date(2026, 7, 1))
        self._line(2, revenue=Decimal("200"), cost=Decimal("100"), invoice_month=None)

        by_out = self._summary("out")
        out_keys = {(r["month"]) for r in by_out["rows"] if r["level"] == "category"}
        self.assertEqual(out_keys, {"2026-06", "2026-07"})

        by_invoice = self._summary("invoice")
        invoice_keys = {(r["month"]) for r in by_invoice["rows"] if r["level"] == "category"}
        self.assertEqual(invoice_keys, {"2026-07", "(未开票)"})
        # 合计不受口径影响
        self.assertEqual(by_out["grand_total"]["revenue"], by_invoice["grand_total"]["revenue"])
        self.assertEqual(by_out["grand_total"]["revenue"], Decimal("1200.00"))

    def test_summary_expenses_follow_row_month(self):
        # 费用按行所属出库月取数：6 月行用 6 月录的费用
        self._line(1, revenue=Decimal("100"), cost=Decimal("10"), out_month=self.june)
        self._line(2, revenue=Decimal("100"), cost=Decimal("10"))
        self._expense(self.june, admin_expense=Decimal("30"))
        self._expense(self.period.month, admin_expense=Decimal("70"))

        summary = self._summary("out")
        by_key = {(r["month"], r["channel"], r["category"]): r for r in summary["rows"]}
        self.assertEqual(
            by_key[("2026-06", "国内销售", "小计")]["admin_expense"], Decimal("30.00")
        )
        self.assertEqual(
            by_key[("2026-07", "国内销售", "小计")]["admin_expense"], Decimal("70.00")
        )
        self.assertEqual(summary["grand_total"]["admin_expense"], Decimal("100.00"))

    # —————————————— 页面与导出 ——————————————

    def test_expense_entry_page_lists_every_month_and_exports(self):
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"), out_month=self.june)
        self._line(2, revenue=Decimal("1000"), cost=Decimal("600"))
        client = self.client
        client.force_login(self.user)
        url = reverse("profit_statement", args=[self.period.id])

        response = client.get(url)
        self.assertEqual(response.status_code, 200)
        for text in (
            "按月费用录入", "税金及附加", "管理费用", "销售费用", "财务费用",
            "营业外收入", "营业外支出", "所得税费用",
            "销售毛利", "营业利润", "利润总额", "净利润", "净利润率",
            "损益汇总表",
        ):
            self.assertContains(response, text)
        html = response.content.decode("utf-8")
        # 表头 = 期间 + 7 项费用；每个月一行（2 个月 → 14 个输入框）
        self.assertIn("<th>期间</th>", html)
        self.assertEqual(html.count('class="expense-input"'), 14)
        self.assertIn("<strong>2026-06</strong>", html)
        self.assertIn("<strong>2026-07</strong>", html)
        self.assertIn("exp_2026-06_admin_expense", html)
        self.assertIn("exp_2026-07_admin_expense", html)
        # 已删掉「期间 从…到…」这一行筛选
        self.assertNotIn("期间 从", html)
        self.assertNotIn('name="mfrom"', html)
        self.assertNotIn('name="mto"', html)

        response = client.post(
            url,
            {
                "basis": "out",
                "exp_2026-06_admin_expense": "600",
                "exp_2026-06_tax_surcharge": "1,000",
                "exp_2026-07_admin_expense": "1000",
                "exp_2026-07_income_tax": "700",
            },
        )
        self.assertRedirects(response, url)
        june = ProfitStatementExpense.objects.get(period=self.period, month=self.june)
        self.assertEqual(june.tax_surcharge, Decimal("1000.00"))  # 千分位可识别
        self.assertEqual(june.admin_expense, Decimal("600.00"))
        self.assertEqual(june.selling_expense, Decimal("0.00"))
        july = ProfitStatementExpense.objects.get(period=self.period, month=self.period.month)
        self.assertEqual(july.admin_expense, Decimal("1000.00"))
        self.assertEqual(july.income_tax, Decimal("700.00"))

        # 开票期间视图可切换
        self.assertEqual(client.get(url + "?basis=invoice").status_code, 200)
        self.assertEqual(client.get(url + "?basis=zzz").status_code, 200)

        export = client.get(reverse("profit_statement_export", args=[self.period.id]))
        self.assertEqual(export.status_code, 200)
        workbook = load_workbook(BytesIO(export.content))
        self.assertEqual(
            workbook.sheetnames, ["损益表", "汇总-按出库期间", "汇总-按开票期间"]
        )
        detail = workbook["损益表"]
        self.assertEqual(
            [detail.cell(1, column).value for column in range(1, len(HEADERS) + 1)],
            list(HEADERS),
        )
        self.assertEqual(detail.max_row, 4)  # 表头 + 2 行 + 合计
        # 明细表最后一行要写明「合计」（用户 2026-09-28）
        self.assertEqual(detail.cell(detail.max_row, 1).value, "合计")
        summary_sheet = workbook["汇总-按出库期间"]
        self.assertEqual(
            [summary_sheet.cell(1, column).value for column in range(1, 4)],
            ["期间", "销售渠道", "销售大类"],
        )
        # 汇总表最后一行是合计，收入 = 2000
        self.assertEqual(summary_sheet.cell(summary_sheet.max_row, 1).value, "合计")
        self.assertEqual(float(summary_sheet.cell(summary_sheet.max_row, 5).value), 2000.0)

    def test_category_summary_columns_and_totals(self):
        # 页面顶端汇总：期间 → 销售类别，8 个指标列，含期间小计与全部合计
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"), category="纱线")
        self._line(2, revenue=Decimal("500"), cost=Decimal("250"), category="面料")
        self._line(3, revenue=Decimal("200"), cost=Decimal("120"),
                   category="纱线", channel="国外销售")
        self._line(4, revenue=Decimal("100"), cost=Decimal("60"),
                   category="面料", channel="国外销售", out_month=self.june)
        self._expense(self.period.month, admin_expense=Decimal("300"))
        self._expense(self.june, admin_expense=Decimal("100"))

        data = build_profit_statement(self.period)
        summary = build_category_summary(data["rows"], "out")

        self.assertEqual(
            list(summary["headers"]),
            ["期间", "销售类别", "销售收入", "销售成本", "销售毛利", "销售毛利率(%)",
             "营业利润", "利润总额", "净利润", "净利润率(%)"],
        )

        by_key = {(r["month"], r["category"]): r for r in summary["rows"]}
        # 7 月纱线 = 两行（国内 1000 + 国外 200）合起来
        yarn = by_key[("2026-07", "纱线")]
        self.assertEqual(yarn["revenue"], Decimal("1200.00"))
        self.assertEqual(yarn["cost"], Decimal("720.00"))
        self.assertEqual(yarn["gross"], Decimal("480.00"))
        self.assertEqual(yarn["gross_margin"], Decimal("40.00"))
        # 7 月管理费用 300 按该月收入占比摊 → 该月合计扣 300
        july = by_key[("2026-07", "小计")]
        self.assertEqual(july["revenue"], Decimal("1700.00"))
        self.assertEqual(july["gross"], Decimal("730.00"))
        self.assertEqual(july["operating_profit"], Decimal("430.00"))
        self.assertEqual(july["total_profit"], Decimal("430.00"))
        self.assertEqual(july["net_profit"], Decimal("430.00"))
        # 6 月
        june = by_key[("2026-06", "小计")]
        self.assertEqual(june["revenue"], Decimal("100.00"))
        self.assertEqual(june["gross"], Decimal("40.00"))
        self.assertEqual(june["operating_profit"], Decimal("-60.00"))

        # 合计 = 全部行合计（含费用），与明细合计一致
        grand = summary["grand_total_row"]
        self.assertEqual(grand["month"], "合计")
        self.assertEqual(grand["revenue"], data["totals"]["revenue"])
        self.assertEqual(grand["cost"], data["totals"]["cost"])
        self.assertEqual(grand["gross"], data["totals"]["gross"])
        self.assertEqual(grand["operating_profit"], data["totals"]["operating_profit"])
        self.assertEqual(grand["net_profit"], data["totals"]["net_profit"])
        # 期间小计之和 = 合计
        month_totals = [r for r in summary["rows"] if r["level"] == "month"]
        self.assertEqual(
            sum((r["revenue"] for r in month_totals), Decimal("0")),
            grand["revenue"],
        )
        # 每期类别行之和 = 该期小计
        for block in summary["months"]:
            self.assertEqual(
                sum((r["revenue"] for r in block["rows"]), Decimal("0")),
                block["total"]["revenue"],
            )
        # 7 月大类按收入倒序：纱线(1200) → 面料(500)
        self.assertEqual(
            [r["category"] for r in summary["rows"]
             if r["level"] == "category" and r["month"] == "2026-07"],
            ["纱线", "面料"],
        )

    def test_category_summary_follows_basis_and_unclassified(self):
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"),
                   out_month=self.june, invoice_month=date(2026, 7, 1))
        self._line(2, revenue=Decimal("200"), cost=Decimal("100"),
                   invoice_month=None, category="")
        data = build_profit_statement(self.period)

        by_invoice = build_category_summary(data["rows"], "invoice")
        keys = {(r["month"], r["category"]) for r in by_invoice["rows"]}
        self.assertIn(("2026-07", "纱线"), keys)
        self.assertIn(("(未开票)", "未分类"), keys)
        # 合计不随口径变化
        by_out = build_category_summary(data["rows"], "out")
        self.assertEqual(
            by_out["grand_total_row"]["revenue"], by_invoice["grand_total_row"]["revenue"]
        )
        self.assertEqual(
            by_invoice["grand_total_row"]["revenue"], Decimal("1200.00")
        )

    def test_category_summary_period_range_filter(self):
        # 区间筛选：只显示所选期间，合计=所选区间合计；不影响别处
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"), category="纱线",
                   out_month=date(2026, 5, 1))
        self._line(2, revenue=Decimal("2000"), cost=Decimal("1200"), category="纱线")
        self._line(3, revenue=Decimal("500"), cost=Decimal("300"), category="面料")
        data = build_profit_statement(self.period)

        # 期间选项：升序、（未开票）最后
        options = [item["value"] for item in summary_period_options(data["rows"], "out")]
        self.assertEqual(options, ["2026-05", "2026-07"])

        only_may = build_category_summary(data["rows"], "out", "2026-05", "2026-05")
        self.assertEqual(
            [r["month"] for r in only_may["rows"] if r["level"] == "month"], ["2026-05"]
        )
        self.assertEqual(only_may["grand_total_row"]["revenue"], Decimal("1000.00"))
        self.assertTrue(only_may["has_range"])

        both = build_category_summary(data["rows"], "out", "2026-05", "2026-07")
        self.assertEqual(
            [r["month"] for r in both["rows"] if r["level"] == "month"],
            ["2026-05", "2026-07"],
        )
        self.assertEqual(both["grand_total_row"]["revenue"], Decimal("3500.00"))

        # 只给「从」不给「到」= 从该期到最新
        from_only = build_category_summary(data["rows"], "out", "2026-07", None)
        self.assertEqual(from_only["grand_total_row"]["revenue"], Decimal("2500.00"))
        # 只给「到」不给「从」= 最早到该期
        to_only = build_category_summary(data["rows"], "out", None, "2026-05")
        self.assertEqual(to_only["grand_total_row"]["revenue"], Decimal("1000.00"))

        # 范围外 → 空表
        empty = build_category_summary(data["rows"], "out", "2026-09", "2026-09")
        self.assertTrue(empty["empty"])
        self.assertEqual(len(empty["rows"]), 1)  # 只剩合计行（0）

    def test_category_summary_range_handles_uninvoiced_and_swapped_bounds(self):
        # 开票口径下的「(未开票)」可被选中；页面里起止反了会自动对调、非法值被忽略
        self._line(1, revenue=Decimal("300"), cost=Decimal("100"),
                   out_month=date(2026, 5, 1), invoice_month=date(2026, 7, 1))
        self._line(2, revenue=Decimal("200"), cost=Decimal("100"), invoice_month=None)
        data = build_profit_statement(self.period)

        options = [item["value"] for item in summary_period_options(data["rows"], "invoice")]
        self.assertEqual(options, ["2026-07", "(未开票)"])
        uninv = build_category_summary(data["rows"], "invoice", "(未开票)", "(未开票)")
        self.assertEqual(uninv["grand_total_row"]["revenue"], Decimal("200.00"))

        client = self.client
        client.force_login(self.user)
        url = reverse("profit_statement", args=[self.period.id])

        # 期间筛选：只勾 2026-05 → 汇总只剩该期间，合计按筛选重算
        response = client.get(url + "?basis=out&m=2026-05")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.context["selected_periods"], ["2026-05"])
        self.assertTrue(response.context["summary"]["has_filter"])
        months = [r["month"] for r in response.context["summary"]["rows"]
                  if r["level"] == "month"]
        self.assertEqual(months, ["2026-05"])
        self.assertEqual(response.context["summary"]["grand_total"]["revenue"],
                         Decimal("300.00"))

        # 一个都不勾 + 提交标记 = 全部期间（「显示全部」）
        response = client.get(url + "?basis=out&f=1")
        self.assertEqual(response.context["selected_periods"], [])
        self.assertFalse(response.context["summary"]["has_filter"])
        months = [r["month"] for r in response.context["summary"]["rows"]
                  if r["level"] == "month"]
        self.assertEqual(months, ["2026-05", "2026-07"])

        # 不存在的期间被忽略 → 视为不筛选
        response = client.get(url + "?basis=out&m=2000-01")
        self.assertEqual(response.context["selected_periods"], [])

        # 开票口径下「(未开票)」可被勾选，且只统计未开票行
        response = client.get(url + "?basis=invoice&m=%28%E6%9C%AA%E5%BC%80%E7%A5%A8%29")
        self.assertEqual(response.context["selected_periods"], ["(未开票)"])
        self.assertEqual(response.context["summary"]["grand_total"]["revenue"],
                         Decimal("200.00"))
        # 切换口径的链接保留期间筛选参数
        html = response.content.decode("utf-8")
        self.assertIn("m=%28%E6%9C%AA%E5%BC%80%E7%A5%A8%29", html)

    def test_category_summary_category_filter(self):
        # 类别筛选：单选、多选、不选（全部）；与期间条件叠加；合计随之变化
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"), category="纱线")
        self._line(2, revenue=Decimal("500"), cost=Decimal("250"), category="面料")
        self._line(3, revenue=Decimal("200"), cost=Decimal("120"), category="纤维",
                   out_month=date(2026, 5, 1))
        data = build_profit_statement(self.period)

        options = [item["value"] for item in summary_category_options(data["rows"])]
        self.assertEqual(options, ["纱线", "面料", "纤维"])  # 按收入倒序

        # 单选
        only_yarn = build_category_summary(data["rows"], "out", categories=["纱线"])
        self.assertEqual([r["category"] for r in only_yarn["rows"] if r["level"] == "category"],
                         ["纱线"])
        self.assertEqual(only_yarn["grand_total_row"]["revenue"], Decimal("1000.00"))
        self.assertTrue(only_yarn["has_category_filter"])

        # 多选
        two = build_category_summary(data["rows"], "out", categories=["纱线", "纤维"])
        self.assertEqual(two["grand_total_row"]["revenue"], Decimal("1200.00"))
        self.assertEqual({r["category"] for r in two["rows"] if r["level"] == "category"},
                         {"纱线", "纤维"})

        # 不选 = 全部
        all_cats = build_category_summary(data["rows"], "out")
        self.assertEqual(all_cats["grand_total_row"]["revenue"], Decimal("1700.00"))
        self.assertFalse(all_cats["has_category_filter"])

        # 与期间叠加：5 月只有纤维
        combo = build_category_summary(data["rows"], "out", "2026-05", "2026-05", ["纤维"])
        self.assertEqual(combo["grand_total_row"]["revenue"], Decimal("200.00"))
        # 叠加后无交集 → 空表
        empty = build_category_summary(data["rows"], "out", "2026-05", "2026-05", ["纱线"])
        self.assertTrue(empty["empty"])

    def test_summary_category_filter_on_page(self):
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"), category="纱线")
        self._line(2, revenue=Decimal("500"), cost=Decimal("250"), category="面料")
        client = self.client
        client.force_login(self.user)
        url = reverse("profit_statement", args=[self.period.id])

        # 页面：三个筛选都是复选框，且选项里不带金额
        response = client.get(url + "?basis=out")
        html = response.content.decode("utf-8")
        self.assertIn('type="checkbox" name="cat"', html)
        self.assertIn('type="checkbox" name="m"', html)
        self.assertIn('type="checkbox" name="ch"', html)
        self.assertIn('value="纱线"', html)
        self.assertIn('value="面料"', html)
        self.assertIsNone(re.search(r"（[\d,]+\.\d\d）", html), "筛选项不应显示金额")
        # 已勾选项回显为 checked
        response = client.get(url + "?basis=out&cat=纱线")
        self.assertIn('value="纱线" checked', response.content.decode("utf-8"))

        # 单选：只勾一个销售大类 → 汇总（含合计）跟着筛
        response = client.get(url + "?basis=out&cat=纱线")
        self.assertEqual(response.context["selected_categories"], ["纱线"])
        self.assertEqual(
            {r["category"] for r in response.context["summary"]["rows"]
             if r["level"] == "category"},
            {"纱线"},
        )
        self.assertEqual(response.context["summary"]["grand_total"]["revenue"],
                         Decimal("1000.00"))
        self.assertTrue(response.context["summary"]["has_filter"])

        # 多选参数生效
        response = client.get(url + "?basis=out&cat=纱线&cat=面料")
        self.assertEqual(response.context["selected_categories"], ["纱线", "面料"])
        cats = {r["category"] for r in response.context["summary"]["rows"]
                if r["level"] == "category"}
        self.assertEqual(cats, {"纱线", "面料"})

        # 不存在的取值被忽略 → 大类视为不筛选
        response = client.get(url + "?basis=out&cat=不存在")
        self.assertEqual(response.context["selected_categories"], [])
        # 链接没带提交标记也没指定期间 → 期间仍默认当月，因此仍算已筛选
        self.assertEqual(response.context["selected_periods"], ["2026-07"])
        self.assertTrue(response.context["summary"]["has_filter"])
        self.assertEqual(response.context["summary"]["grand_total"]["revenue"],
                         Decimal("1500.00"))

        # 口径切换链接与「显示全部」保留当前筛选（URL 编码）
        response = client.get(url + "?basis=out&cat=纱线&ch=国内销售")
        html = response.content.decode("utf-8")
        self.assertIn("cat=%E7%BA%B1%E7%BA%BF", html)  # “纱线”的 URL 编码
        self.assertIn("ch=%E5%9B%BD%E5%86%85%E9%94%80%E5%94%AE", html)  # “国内销售”的 URL 编码
        self.assertIn("显示全部", html)

    def test_expense_entry_lists_all_ledger_months(self):
        # 台账里出现过的每个月都要列出来（不再有「期间 从…到…」筛选）
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"),
                   out_month=date(2026, 5, 1))
        self._line(2, revenue=Decimal("1000"), cost=Decimal("600"), out_month=self.june)
        self._line(3, revenue=Decimal("1000"), cost=Decimal("600"))
        self._expense(self.june, admin_expense=Decimal("100"))
        client = self.client
        client.force_login(self.user)
        url = reverse("profit_statement", args=[self.period.id])

        html = client.get(url).content.decode("utf-8")
        for label in ("2026-05", "2026-06", "2026-07"):
            self.assertIn("<strong>%s</strong>" % label, html)
        self.assertEqual(html.count('class="expense-input"'), 21)  # 3 个月 × 7 项
        # 已录入的月份有标记，且已录金额回填在输入框里
        self.assertIn("已录入", html)
        self.assertIn('value="100.00"', html)
        self.assertNotIn("期间 从", html)
        # 台账里没有的月份即使提交也不会写库
        client.post(url, {"exp_2026-04_admin_expense": "10"})
        self.assertFalse(
            ProfitStatementExpense.objects.filter(
                period=self.period, month=date(2026, 4, 1)
            ).exists()
        )

    def test_page_export_sits_next_to_the_print_button(self):
        """用户 2026-09-28：整份损益表的「导出Excel」从页头挪到「按月费用录入」标题行，
        与（由 base.html 注入的）「打印本表」放在一起；页头不再有导出按钮。"""
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"))
        client = self.client
        client.force_login(self.user)
        html = client.get(
            reverse("profit_statement", args=[self.period.id])
        ).content.decode("utf-8")
        page_head = html.split("</section>", 1)[0]
        self.assertNotIn("导出Excel", page_head)
        seg = html[html.index("<h2>按月费用录入</h2>"):]
        seg = seg[: seg.index("</table>")]
        self.assertIn('class="heading-actions"', seg)
        self.assertIn(reverse("profit_statement_export", args=[self.period.id]), seg)
        self.assertLess(seg.index('class="heading-actions"'), seg.index("导出Excel"))

    def test_expense_entry_table_has_total_row(self):
        """用户 2026-09-28：「按月费用录入」表最后加一行合计（7 项费用纵向合计）。"""
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"), out_month=self.june)
        self._line(2, revenue=Decimal("2000"), cost=Decimal("1200"))
        self._expense(self.june, tax_surcharge=Decimal("100"), admin_expense=Decimal("50"))
        self._expense(self.period.month, tax_surcharge=Decimal("300"),
                      admin_expense=Decimal("20"))
        client = self.client
        client.force_login(self.user)
        response = client.get(reverse("profit_statement", args=[self.period.id]))
        self.assertEqual(response.status_code, 200)
        html = response.content.decode("utf-8")

        # 合计值 = 两个月相加
        totals = response.context["months_total"]
        self.assertEqual(totals["month_count"], 2)
        values = {item["field"]: item["value"] for item in totals["fields"]}
        self.assertEqual(values["tax_surcharge"], Decimal("400.00"))
        self.assertEqual(values["admin_expense"], Decimal("70.00"))
        self.assertEqual(values["selling_expense"], Decimal("0.00"))
        self.assertEqual(totals["revenue"], Decimal("3000.00"))

        # 页面上确实渲染在表尾（tfoot），且没有多出输入框
        section = html[html.index("<h2>按月费用录入</h2>"):]
        section = section[: section.index("</table>")]
        self.assertIn("<tfoot>", section)
        self.assertIn('class="expense-total-row"', section)
        self.assertIn("<strong>合计</strong>", section)
        # 用户 2026-09-28：合计行不再带「（N 个月）」与「终端销售收入 … · 台账 … 行」小字
        total_row = section.split('class="expense-total-row"', 1)[1]
        self.assertNotIn("个月）", total_row.split("</tr>", 1)[0])
        self.assertNotIn("台账", total_row.split("</tr>", 1)[0])
        self.assertEqual(html.count('class="expense-input"'), 14)   # 2 个月 × 7 项
        # 每行末尾的「合计」列＝该月 7 项费用之和（用户 2026-09-28）
        self.assertIn("<th class=\"num\">合计</th>", section)
        totals = {item["month"]: item["total"] for item in response.context["months"]}
        self.assertEqual(totals[date(2026, 6, 1)], Decimal("150.00"))
        self.assertEqual(totals[date(2026, 7, 1)], Decimal("320.00"))
        self.assertEqual(response.context["months_total"]["grand"], Decimal("470.00"))
        self.assertIn("<th class=\"num\">合计</th>", section)

    def test_profit_table_has_scroll_wrapper_and_sticky_header(self):
        # 损益明细表：外层是可滚动的容器，配 sticky 表头样式（表头固定）＋紧凑列宽
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"), category="纱线")
        client = self.client
        client.force_login(self.user)
        url = reverse("profit_statement", args=[self.period.id])
        html = client.get(url).content.decode("utf-8")
        self.assertIn('class="table-scroll profit-table-wrap"', html)
        self.assertIn('class="data-table profit-table"', html)
        self.assertIn("损益明细表", html)
        # 损益明细表不再依赖内联的横向滚动
        section = html[html.index("损益明细表"):]
        section = section[:section.index("</section>")]
        self.assertNotIn('style="overflow-x:auto;"', section)
        self.assertIn("profit-table-wrap", section)

    def test_summary_is_single_table_with_filters(self):
        # 汇总只有一张表：期间 × 销售渠道 × 销售大类，上方是 期间/销售渠道/销售大类 三个筛选
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"), category="纱线")
        client = self.client
        client.force_login(self.user)
        url = reverse("profit_statement", args=[self.period.id])
        response = client.get(url)
        self.assertEqual(response.status_code, 200)
        html = response.content.decode("utf-8")
        # 汇总表标题改成「损益汇总表」（细分结构写在标题下方的说明里）
        self.assertIn("<h2>损益汇总表</h2>", html)
        # 旧的两张表都不再出现
        self.assertNotIn("月度 × 销售渠道", html)
        self.assertNotIn("汇总：期间 × 销售类别", html)
        # 合并表的表头 = 期间 + 销售渠道 + 销售大类
        for header in ("<th>期间</th>", "<th>销售渠道</th>", "<th>销售大类</th>"):
            self.assertIn(header, html)
        # 三个筛选字段（复选框）
        for name in ('type="checkbox" name="m"', 'type="checkbox" name="ch"',
                     'type="checkbox" name="cat"'):
            self.assertIn(name, html)
        # 用户 2026-09-28：汇总表去掉「按出库期间」口径切换（固定按开票期间），
        # 筛选行加「导出Excel」，与 期间 / 销售渠道 / 销售大类 同一行对齐
        self.assertNotIn("basis=out", html)
        self.assertIn("导出Excel", html)
        self.assertIn(reverse("profit_summary_export", args=[self.period.id]), html)
        # 汇总表在「按月费用录入」之后（原顶端表已删除）
        self.assertLess(html.index("按月费用录入"), html.index("<h2>损益汇总表</h2>"))

    def test_summary_export_follows_filters(self):
        # 损益汇总表筛选行的「导出Excel」：导出的是当前筛选条件下这一张汇总表（按开票期间）
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"), category="纱线")
        self._line(2, revenue=Decimal("500"), cost=Decimal("250"), category="面料")
        client = self.client
        client.force_login(self.user)
        url = reverse("profit_summary_export", args=[self.period.id])

        response = client.get(url)
        self.assertEqual(response.status_code, 200)
        workbook = load_workbook(BytesIO(response.content))
        self.assertEqual(workbook.sheetnames, ["损益汇总表"])
        sheet = workbook["损益汇总表"]
        headers = [cell.value for cell in sheet[1]]
        self.assertEqual(headers[:3], ["期间", "销售渠道", "销售大类"])
        revenue_col = headers.index("终端销售收入")
        rows = list(sheet.iter_rows(values_only=True))
        self.assertEqual(rows[-1][0], "合计")
        self.assertEqual(Decimal(str(rows[-1][revenue_col])), Decimal("1500.00"))

        # 带上「销售大类 = 纱线」筛选：只导这一类
        response = client.get(url + "?cat=%E7%BA%B1%E7%BA%BF")
        workbook = load_workbook(BytesIO(response.content))
        rows = list(workbook["损益汇总表"].iter_rows(values_only=True))
        self.assertEqual(Decimal(str(rows[-1][revenue_col])), Decimal("1000.00"))

    def test_removed_help_texts_are_gone(self):
        # 用户 2026-09-28 要求删掉的四段说明文字，页面上不应再出现
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"))
        client = self.client
        client.force_login(self.user)
        html = client.get(
            reverse("profit_statement", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertNotIn("已按月末加权平均价回填", html)
        self.assertNotIn("每个期间一行", html)
        self.assertNotIn("结构：期间 × 销售渠道 × 销售大类", html)
        self.assertNotIn("可按期间/渠道/大类筛选，打开页面默认只看当月", html)
        # 右上角的「报表导出中心」按钮也删掉了（用户 2026-09-28）
        self.assertNotIn("报表导出中心", html)
        self.assertIn(reverse("profit_statement_export", args=[self.period.id]), html)

    def test_summary_row_puts_export_next_to_print(self):
        # 损益汇总表：导出Excel 与 打印本表 挨在一起（打印按钮注入到同一行的 .print-here）
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"))
        client = self.client
        client.force_login(self.user)
        html = client.get(
            reverse("profit_statement", args=[self.period.id])
        ).content.decode("utf-8")
        section = html[html.index("<h2>损益汇总表</h2>"):]
        section = section[: section.index("</table>")]
        form = section.split("<form", 1)[1].split("</form>", 1)[0]
        self.assertIn('class="filter-grid compact-filters"', form)
        self.assertIn('class="filter-actions"', form)
        # 顺序：确定 -> 导出Excel -> 打印本表（注入位）
        actions = form.split('class="filter-actions"', 1)[1]
        self.assertIn("导出Excel", actions)
        self.assertIn('class="print-here"', actions)
        self.assertLess(actions.index("导出Excel"), actions.index('class="print-here"'))
        self.assertIn(reverse("profit_summary_export", args=[self.period.id]), actions)

    def test_summary_default_shows_current_period_month(self):
        # 打开页面（不带任何筛选）默认只看当月：期间勾当前期间月份，合计按当月重算
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"))
        self._line(2, revenue=Decimal("200"), cost=Decimal("100"), out_month=self.june,
                   invoice_month=self.june)
        client = self.client
        client.force_login(self.user)
        url = reverse("profit_statement", args=[self.period.id])

        response = client.get(url)
        self.assertEqual(response.context["selected_periods"], ["2026-07"])
        self.assertTrue(response.context["default_period_applied"])
        self.assertTrue(response.context["summary"]["has_filter"])
        self.assertEqual(
            [r["month"] for r in response.context["summary"]["rows"]
             if r["level"] == "month"],
            ["2026-07"],
        )
        self.assertEqual(response.context["summary"]["grand_total"]["revenue"],
                         Decimal("1000.00"))
        html = response.content.decode("utf-8")
        self.assertIn("<h2>\u635f\u76ca\u6c47\u603b\u8868</h2>", html)
        self.assertIn('value="2026-07" checked', html)

        # 「显示全部」带提交标记 f=1 → 三个筛选都清空，回到全部期间
        self.assertIn("f=1", html)
        response = client.get(url + "?basis=out&f=1")
        self.assertEqual(response.context["selected_periods"], [])
        self.assertFalse(response.context["default_period_applied"])
        self.assertEqual(response.context["summary"]["grand_total"]["revenue"],
                         Decimal("1200.00"))

    def test_summary_default_period_falls_back_to_latest_ledger_month(self):
        # 台账里没有当前期间那个月 → 默认退到最近有数据的期间（不留空表）
        self._line(1, revenue=Decimal("200"), cost=Decimal("100"),
                   out_month=date(2026, 5, 1), invoice_month=date(2026, 5, 1))
        self._line(2, revenue=Decimal("1000"), cost=Decimal("600"), out_month=self.june,
                   invoice_month=self.june)
        client = self.client
        client.force_login(self.user)
        url = reverse("profit_statement", args=[self.period.id])
        response = client.get(url)
        self.assertEqual(response.context["selected_periods"], ["2026-06"])
        self.assertEqual(response.context["summary"]["grand_total"]["revenue"],
                         Decimal("1000.00"))

    def test_summary_default_period_on_invoice_basis(self):
        # 开票口径：默认仍是当前期间月份；台账全未开票（没有该月份）→ 默认落到「(未开票)」
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"))
        self._line(2, revenue=Decimal("200"), cost=Decimal("100"),
                   invoice_month=None)
        client = self.client
        client.force_login(self.user)
        url = reverse("profit_statement", args=[self.period.id])

        response = client.get(url + "?basis=invoice")
        self.assertEqual(response.context["selected_periods"], ["2026-07"])
        self.assertEqual(response.context["summary"]["grand_total"]["revenue"],
                         Decimal("1000.00"))

        SalesLedgerLine.objects.filter(row_no=1).update(invoice_month=None)
        response = client.get(url + "?basis=invoice")
        self.assertEqual(response.context["selected_periods"], ["(\u672a\u5f00\u7968)"])
        self.assertEqual(response.context["summary"]["grand_total"]["revenue"],
                         Decimal("1200.00"))

    def test_summary_table_has_scroll_wrapper_and_sticky_header(self):
        # 损益汇总表：外层是带滚动条的容器（横竖都能滚），滚动时表头固定
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"))
        client = self.client
        client.force_login(self.user)
        url = reverse("profit_statement", args=[self.period.id])
        html = client.get(url).content.decode("utf-8")
        self.assertIn('class="table-scroll summary-table-wrap"', html)
        self.assertIn('class="data-table summary-table"', html)
        section = html[html.index("<h2>\u635f\u76ca\u6c47\u603b\u8868</h2>"):]
        section = section[:section.index("</section>")]
        # 汇总表自己带滚动容器，不再依赖内联的横向滚动
        self.assertNotIn('style="overflow-x:auto;"', section)
        self.assertIn("summary-table-wrap", section)

    # ──────────────── 2026-09-15 用户要求的四项调整 ────────────────

    def test_expense_input_shows_thousands_separator(self):
        # 按月费用录入：金额带分节号显示；带分节号的值能原样回写
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"))
        self._expense(self.period.month, admin_expense=Decimal("2098509.12"))
        client = self.client
        client.force_login(self.user)
        url = reverse("profit_statement", args=[self.period.id])
        html = client.get(url).content.decode("utf-8")
        self.assertIn('value="2,098,509.12"', html)
        self.assertIn('inputmode="decimal"', html)
        # 不能用 type=number：浏览器会把带逗号的值当非法值清空
        cell = html[html.index("exp_2026-07_admin_expense"):]
        cell = cell[: cell.index(">") + 1]
        self.assertIn('type="text"', cell)
        self.assertNotIn('type="number"', html)

        # 原样回写：千分位 / 半角逗号 / 普通数字都能存对
        response = client.post(url, {
            "exp_2026-07_admin_expense": "2,098,509.12",
            "exp_2026-07_selling_expense": "1,234.5",
            "exp_2026-07_finance_expense": "987",
        })
        self.assertRedirects(response, url)
        saved = ProfitStatementExpense.objects.get(
            period=self.period, month=self.period.month
        )
        self.assertEqual(saved.admin_expense, Decimal("2098509.12"))
        self.assertEqual(saved.selling_expense, Decimal("1234.50"))
        self.assertEqual(saved.finance_expense, Decimal("987.00"))
        # 回显仍然是分节号
        html = client.get(url).content.decode("utf-8")
        self.assertIn('value="2,098,509.12"', html)
        self.assertIn('value="1,234.50"', html)

    def test_summary_table_has_grid_lines(self):
        # 损益汇总表加网格线（行列分隔清楚）
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"))
        client = self.client
        client.force_login(self.user)
        html = client.get(
            reverse("profit_statement", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertIn('class="data-table summary-table"', html)
        css = (settings.BASE_DIR / "static" / "settlement" / "app.css").read_text(
            encoding="utf-8"
        )
        block = re.search(r"\.summary-table th, \.summary-table td\s*\{([^}]*)\}", css)
        self.assertIsNotNone(block, "缺少汇总表单元格样式")
        self.assertIn("padding", block.group(1))
        # 网格线走全局规则，汇总表也覆盖到
        self.assertRegex(css, r"th, td \{[^}]*border: 1px solid #e3e9ed")

    def test_category_parenthetical_shows_short_name(self):
        # 销售大类带明细括注的只显示主体；括注明细写在表下说明里
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"),
                   category="其他成品（手套、服装、床垫等）")
        self._line(2, revenue=Decimal("500"), cost=Decimal("250"), category="纱线")
        data = build_profit_statement(self.period)

        # 明细行与筛选选项都用主体名
        cats = {row["category"] for row in data["rows"]}
        self.assertIn("其他成品", cats)
        self.assertNotIn("其他成品（手套、服装、床垫等）", cats)
        options = [item["value"] for item in summary_category_options(data["rows"])]
        self.assertEqual(options, ["其他成品", "纱线"])  # 按收入倒序（其他成品 1000 > 纱线 500）
        # 带不同括注的同一主体归到一起
        self.assertEqual(summary_category_options(data["rows"])[1]["row_count"], 1)

        client = self.client
        client.force_login(self.user)
        url = reverse("profit_statement", args=[self.period.id])
        html = client.get(url).content.decode("utf-8")
        summary_section = html[html.index("<h2>损益汇总表</h2>"):]
        summary_section = summary_section[: summary_section.index("</section>")]
        detail_section = html[html.index("<h2>损益明细表"):]
        detail_section = detail_section[: detail_section.index("</section>")]
        for section in (summary_section, detail_section):
            self.assertIn("<td>其他成品</td>", section)
            self.assertNotIn("其他成品（手套", section)
        # 括注明细落在汇总表下方的说明里
        self.assertIn("（手套、袜子、服装、床垫等）", summary_section)
        self.assertIn("手套、袜子、服装、床垫", summary_section)
        # 筛选用主体名，合计跟着筛
        response = client.get(url + "?basis=out&cat=其他成品")
        self.assertEqual(response.context["selected_categories"], ["其他成品"])
        self.assertEqual(response.context["summary"]["grand_total"]["revenue"],
                         Decimal("1000.00"))

    def test_category_label_helper(self):
        # 归类显示规则：去掉结尾括注；空值归未分类；无括注原样保留
        self.assertEqual(category_label("其他成品（手套、服装、床垫等）"), "其他成品")
        self.assertEqual(category_label("其他成品(手套)"), "其他成品")
        self.assertEqual(category_label("纱线"), "纱线")
        self.assertEqual(category_label("  "), "未分类")
        self.assertEqual(category_label("（全括注）"), "（全括注）")  # 全被吃掉时保留原名

    def test_detail_section_renamed_to_detail_table(self):
        # 「逐行损益」改名为「损益明细表」；用户 2026-09-28：标题只留表名，
        # 不再带「（共 N 行，本页 N 行）」、不再有「台账合计…」小字与「合计行按当前筛选计算」标签
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"))
        client = self.client
        client.force_login(self.user)
        html = client.get(
            reverse("profit_statement", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertIn("<h2>损益明细表</h2>", html)
        self.assertNotIn("逐行损益", html)
        self.assertNotIn("本页", html)
        self.assertNotIn("台账合计", html)
        self.assertNotIn("合计行按当前筛选计算", html)

    def test_locked_period_keeps_expenses_unchanged(self):
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"))
        self._expense(self.period.month, admin_expense=Decimal("100"))
        self.period.status = AccountingPeriod.Status.LOCKED
        self.period.save(update_fields=["status"])

        client = self.client
        client.force_login(self.user)
        url = reverse("profit_statement", args=[self.period.id])
        response = client.post(url, {"exp_2026-07_admin_expense": "999"})
        self.assertRedirects(response, url)
        self.assertEqual(
            ProfitStatementExpense.objects.get(
                period=self.period, month=self.period.month
            ).admin_expense,
            Decimal("100.00"),
        )


class TableScrollContainerTests(SettlementFixtureMixin, TestCase):
    """各页面业务板块：超出一屏的表格统一放进可滚动容器。

    口径（用户 2026-09-15 要求）：其他页面的业务板块超出一个屏幕的也加滚动条。
    实现：`.table-wrap` / `.table-scroll` 统一为滚动容器（横竖滚动 + 滚动时表头固定），
    分页控件留在容器外，保证始终可见；两行表头的表单独给出错位 top 偏移。
    """

    #: 用统一样式的页面（表格容器 = table-wrap / table-scroll）
    WRAPPED_PAGES = (
        "inventory_ledger", "inventory_matches", "inventory_results",
        "purchase_contracts", "purchase_matches", "purchase_receipts",
        "purchase_accruals", "ending_wip",
        "sales_uninvoiced", "reconciliation",
        "import_batches", "work_order_costs",
        "work_order_report", "work_order_imports", "work_report_imports",
        "material_issue_imports", "bom_imports", "order_completion_report",
        "order_material_report", "price_table",
        "workshop_settlements", "production_cost_summary",
    )

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="scroll-user", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 8, 1))
        self.client.force_login(self.user)

    def _html(self, url_name, *args):
        resp = self.client.get(reverse(url_name, args=list(args) or [self.period.id]))
        self.assertEqual(resp.status_code, 200, url_name)
        return resp.content.decode("utf-8")

    def test_every_page_wraps_tables_in_scroll_container(self):
        # 每张表都要在 table-wrap / table-scroll 里（否则不会出现滚动条）
        for name in self.WRAPPED_PAGES:
            html = self._html(name)
            if "<table" not in html:
                continue
            wraps = html.count('class="table-wrap"') + html.count('class="table-scroll')
            tables = html.count("<table")
            self.assertGreaterEqual(wraps, tables, "%s: 有表格没进滚动容器" % name)

    def test_no_inline_horizontal_overflow_left(self):
        # 统一交给样式表，模板里不再留内联的横向滚动
        for name in ("production_cost_summary", "profit_statement", "price_table"):
            self.assertNotIn('style="overflow-x:auto;"', self._html(name), name)

    def test_shared_scroll_styles_cover_both_wrappers(self):
        css = (settings.BASE_DIR / "static" / "settlement" / "app.css").read_text(
            encoding="utf-8"
        )
        block = re.search(r"\.table-wrap,\s*\.table-scroll\s*\{([^}]*)\}", css)
        self.assertIsNotNone(block, "缺少 .table-wrap/.table-scroll 统一规则")
        body = block.group(1)
        self.assertIn("max-height", body)      # 高度上限 → 纵向滚动条
        self.assertIn("overflow: auto", body)  # 横竖都能滚
        self.assertIn("border-radius", body)
        # 滚动时表头固定
        head = re.search(r"\.table-wrap thead th,\s*\.table-scroll thead th\s*\{([^}]*)\}", css)
        self.assertIsNotNone(head, "缺少表头 sticky 规则")
        self.assertIn("position: sticky", head.group(1))
        self.assertIn("top: 0", head.group(1))
        # 长列表（期间详情「最近导入」）也要能在卡片内滚动
        self.assertRegex(css, r"\.compact-list\s*\{[^}]*max-height[^}]*overflow: auto")
        # 行高 25 磅（≈33.3px）：全站表格一行的高度统一
        self.assertRegex(css, r"th, td \{[^}]*height: 25pt")

    def test_price_table_two_row_header_offsets(self):
        # 单价表：两行表头，第一行贴顶、第二行贴在其下（28px），不互相盖住
        html = self._html("price_table")
        self.assertIn('class="price-table"', html)
        self.assertIn(".price-table thead tr:first-child th { top: 0;", html)
        self.assertIn(".price-table thead tr:nth-child(2) th { top: 30px; }", html)
        self.assertIn('class="table-wrap"', html)

    def test_price_table_has_field_filters(self):
        # 单价表（用户 2026-09-30）：恢复字段筛选 —— 期间范围 / 物料编码 / 类别，
        # 并且「导出Excel」要和「打印本表」并排（打印由 base.html 注入到 .print-here）
        html = self._html("price_table")
        self.assertIn('name="p_from"', html)       # 期间（从）
        self.assertIn('name="p_to"', html)         # 期间（到）
        self.assertIn('name="code"', html)         # 物料编码（包含）
        self.assertIn('name="category"', html)     # 类别
        self.assertIn("print-here", html)          # 导出按钮的落点（打印本表挨着它）
        self.assertIn("?export=1", html)
        # 表格本体与表头「类别」列不变
        self.assertIn('class="price-table"', html)
        self.assertIn('<th rowspan="2">类别</th>', html)

    def test_total_cost_two_row_header_offsets(self):
        # 成本构成表（原「成本项目×四大块」，已并入生产成本页）：同样是两行表头，第二行贴在第一行下方
        html = self._html("production_cost_summary")
        self.assertIn('class="table-scroll"', html)
        self.assertIn(".total-cost-table thead tr:nth-child(2) th { top: 30px; }", html)
        self.assertIn(".total-cost-table thead th[rowspan] { vertical-align: top; }", html)

    def test_pagination_stays_outside_scroll_container(self):
        # 分页控件在滚动容器之外 → 表格在容器内滚动时，翻页按钮始终可见
        # （分页只在多于一页时渲染，所以按模板结构校验）
        tpl_dir = settings.BASE_DIR / "templates" / "settlement"
        checked = 0
        for path in sorted(tpl_dir.glob("*.html")):
            text_ = path.read_text(encoding="utf-8")
            if "_pagination.html" not in text_:
                continue
            before = text_[: text_.index("_pagination.html")]
            open_idx = before.rfind('<div class="table-wrap"')
            if open_idx < 0:
                continue
            self.assertGreater(
                before.rfind("</div>"), open_idx,
                "%s: 分页被包进滚动容器里了" % path.name,
            )
            checked += 1
        self.assertGreaterEqual(checked, 10)  # 至少覆盖十个带分页的列表页

    def test_profit_expense_table_uses_shared_container(self):
        # 按月费用录入的表也用统一样式（空数据下该表不渲染，故按模板结构校验）
        path = settings.BASE_DIR / "templates" / "settlement" / "profit_statement.html"
        tpl = path.read_text(encoding="utf-8")
        self.assertIn('<div class="table-scroll">', tpl)
        self.assertNotIn("overflow-x:auto", tpl)
        # 页面上确实存在统一容器（损益汇总表 / 损益明细表，都带额外定位类）
        html = self._html("profit_statement")
        self.assertIn('class="table-scroll', html)
        self.assertNotIn('style="overflow-x:auto;"', html)


class MaterialNameColumnTests(SettlementFixtureMixin, TestCase):
    """物料/产品的编码与名称必须分两列（2026-09-15 用户要求）。

    背景：页面原来把「编码 + 名称」挤在一格里（编码加粗、名称小字），
    而导出 Excel 一直是「物料编码 / 物料名称」两列；现在页面与导出口径一致。
    """

    #: 需要分两列的模板（本次改动涉及的全部页面）
    TEMPLATES = (
        "bom_imports.html", "inventory_ledger.html", "inventory_matches.html",
        "inventory_results.html", "material_issue_imports.html",
        "order_completion_report.html", "order_material_report.html",
        "production_cost_summary.html", "purchase_accruals.html",
        "purchase_contracts.html", "purchase_receipts.html", "ending_wip.html",
        "purchase_matches.html",
        "sales_uninvoiced.html", "work_order_costs.html",
        "work_report_imports.html",
        "source_document_detail.html",
    )

    #: 表头**不在模板里写死**的模板：列名整列照上传的文件（如「工单报工」＝上传的报工记录
    #: 原样 + 6 个计算列），所以这个字面量检查看不到 <th>。
    #: 它们的「编码列必有名称列」改在渲染结果上验（见 OrderCompletionReportPageTests
    #: ::test_detail_table_is_uploaded_report_plus_six_columns），
    #: 同时这里断言它们确实是动态表头，避免豁免被误用。
    DYNAMIC_HEADER_TEMPLATES = {
        "order_completion_report.html": "{% for column in columns %}",
    }

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="material-col-tester", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 9, 1))
        self.batch = self.create_batch(self.period, ImportBatch.Kind.SALES_LEDGER, "cols")
        self.material = Material.objects.create(
            code="RM-COL", name="分列测试物料", uom="kg"
        )

    def test_no_template_merges_code_and_name_in_one_cell(self):
        # 不允许再出现「<strong>编码</strong><small>名称</small>」这种挤在一格的写法
        tpl_dir = settings.BASE_DIR / "templates" / "settlement"
        pattern = re.compile(
            r"<td[^>]*>\s*<strong>\{\{[^}]*\b(material|product|sub_item)_code[^}]*\}\}"
            r"</strong>\s*<small>\{\{[^}]*\1_name",
            re.S,
        )
        offenders = []
        for path in sorted(tpl_dir.glob("*.html")):
            text = path.read_text(encoding="utf-8")
            for match in pattern.finditer(text):
                offenders.append("%s: %s" % (path.name, match.group(0)[:90]))
        self.assertEqual(offenders, [], "以下地方仍把编码和名称挤在一列")

    def test_every_code_column_has_matching_name_column(self):
        # 有「X编码」表头的地方，必须有对应的「X名称」表头
        tpl_dir = settings.BASE_DIR / "templates" / "settlement"
        for name in self.TEMPLATES:
            text = (tpl_dir / name).read_text(encoding="utf-8")
            if name in self.DYNAMIC_HEADER_TEMPLATES:
                # 列名整列照上传文件生成（模板里没有写死的 <th>）→ 这个字面量检查不适用，
                # 用渲染结果验证编码/名称成对（见 DYNAMIC_HEADER_TEMPLATES 的说明）
                self.assertIn(
                    self.DYNAMIC_HEADER_TEMPLATES[name], text,
                    "%s：声称是动态表头，却没看到生成表头的写法" % name,
                )
                continue
            code_headers = set(re.findall(r"<th[^>]*>([^<]*编码)</th>", text))
            self.assertTrue(code_headers, "%s 没有编码列表头" % name)
            for code_header in code_headers:
                name_header = code_header.replace("编码", "名称")
                self.assertIn(
                    ">%s</th>" % name_header, text,
                    "%s 有「%s」却没有「%s」" % (name, code_header, name_header),
                )

    def test_inventory_ledger_page_shows_two_columns(self):
        # 库存流水页：表头两列 + 单元格拆成两个 td
        self.create_txn(
            self.period, self.batch,
            txn_date=date(2026, 9, 10), document_no="IT-001",
            business_type="采购入库", direction="in", material=self.material,
            quantity=Decimal("10"),
        )
        self.client.force_login(self.user)
        html = self.client.get(
            reverse("inventory_ledger", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertIn("<th>物料编码</th>", html)
        self.assertIn("<th>物料名称</th>", html)
        self.assertIn("<td><strong>RM-COL</strong></td>", html)   # 编码列
        self.assertIn("<td>分列测试物料</td>", html)             # 名称列（独立一格）
        # 编码与名称不再是同一个单元格
        self.assertNotIn("<td><strong>RM-COL</strong><small>分列测试物料</small></td>", html)

    def test_sales_uninvoiced_page_shows_two_columns(self):
        # 月末未开票页（表头写「产品」）：产品编码 / 产品名称 两列
        SalesLedgerLine.objects.create(
            period=self.period, batch=self.batch, row_no=1,
            out_month=date(2026, 9, 1),
            invoice_status=SalesLedgerLine.InvoiceStatus.UNINVOICED,
            channel="国内销售", company="测试公司", order_no="SO-1",
            customer_code="C1", customer_name="客户甲",
            sales_category="其他成品", material=self.material,
            material_code=self.material.code, material_name=self.material.name,
            uom="kg", quantity=Decimal("5"), terminal_amount=Decimal("100"),
        )
        self.client.force_login(self.user)
        html = self.client.get(
            reverse("sales_uninvoiced", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertIn("<th>产品编码</th>", html)
        self.assertIn("<th>产品名称</th>", html)
        # 数据行：编码与名称各自独立一格（该页无匹配数据，故按模板结构校验）
        tpl = (settings.BASE_DIR / "templates" / "settlement"
               / "sales_uninvoiced.html").read_text(encoding="utf-8")
        self.assertIn("<td><strong>{{ row.material_code }}</strong></td>", tpl)
        self.assertIn("<td>{{ row.material_name }}</td>", tpl)

    def test_page_column_count_matches_header(self):
        # 拆列后不能错位：每行（按 colspan 累加）的列数 = 表头列数
        self.create_txn(
            self.period, self.batch,
            txn_date=date(2026, 9, 10), document_no="IT-002",
            business_type="采购入库", direction="in", material=self.material,
            quantity=Decimal("10"),
        )
        self.client.force_login(self.user)
        html = self.client.get(
            reverse("inventory_ledger", args=[self.period.id])
        ).content.decode("utf-8")
        table = html[html.index("<table"):html.index("</table>")]
        head = table[table.index("<thead"):table.index("</thead>")]
        head_cols = len(re.findall(r"<th\b", head))
        for row_html in re.findall(r"<tr\b[^>]*>(.*?)</tr>", table, re.S):
            if "<th" in row_html:
                continue
            cells = re.findall(r"<td\b[^>]*>", row_html)
            if not cells:
                continue
            total = 0
            for cell in cells:
                span = re.search(r'colspan="(\d+)"', cell)
                total += int(span.group(1)) if span else 1
            self.assertEqual(total, head_cols, "行列表数 %d != 表头 %d" % (total, head_cols))


class RowHeightTests(SettlementFixtureMixin, TestCase):
    """全站表格行高 = 25 磅（≈33.3px）。"""

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="row-height-tester", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 9, 1))
        self.batch = self.create_batch(self.period, ImportBatch.Kind.SALES_LEDGER, "rowh")
        self.material = Material.objects.create(code="RH-1", name="行高测试物料", uom="kg")

    def test_css_declares_25pt_row_height(self):
        css = (settings.BASE_DIR / "static" / "settlement" / "app.css").read_text(
            encoding="utf-8"
        )
        rule = re.search(r"th, td \{([^}]*)\}", css)
        self.assertIsNotNone(rule)
        self.assertIn("height: 25pt", rule.group(1))
        # 只在一个地方声明行高，避免各页面自己调
        self.assertEqual(css.count("height: 25pt"), 1)

    def test_single_line_row_is_about_25pt(self):
        # 单行内容：行高应贴合 25 磅（33.3px），允许 1px 误差
        self.create_txn(
            self.period, self.batch,
            txn_date=date(2026, 9, 10), document_no="RH-001",
            business_type="采购入库", direction="in", material=self.material,
            quantity=Decimal("1"),
        )
        self.client.force_login(self.user)
        html = self.client.get(
            reverse("inventory_ledger", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertIn("height: 25pt", (
            settings.BASE_DIR / "static" / "settlement" / "app.css"
        ).read_text(encoding="utf-8"))
        # 页面引用的样式表版本号必须和 base.html 里一致（避免改了样式却没换版本号，
        # 浏览器继续吃旧缓存）
        base_tpl = (settings.BASE_DIR / "templates" / "base.html").read_text(
            encoding="utf-8"
        )
        version = re.search(r"app\.css' %\}\?v=([\w\-]+)", base_tpl).group(1)
        self.assertIn("v=" + version, html)


class ProfitDetailFilterTests(SettlementFixtureMixin, TestCase):
    """损益明细表筛选：默认当月 + 出库期间 / 开票期间区间 / 客户名称 / 销售大类 / 产品编码。"""

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="detail-filter-tester", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.batch = self.create_batch(self.period, ImportBatch.Kind.SALES_LEDGER, "df")
        self.june = date(2026, 6, 1)
        # 4 行：6 月两行、7 月两行；客户/大类/编码各不相同，便于逐项筛选
        self.line1 = self._line(1, out_month=self.june, invoice_month=self.june,
                                customer="客户甲", category="纱线", code="11.01",
                                revenue=Decimal("100"), cost=Decimal("60"))
        self.line2 = self._line(2, out_month=self.june, invoice_month=self.period.month,
                                customer="客户乙", category="面料", code="11.02",
                                revenue=Decimal("200"), cost=Decimal("120"))
        self.line3 = self._line(3, invoice_month=None, customer="客户甲", category="纤维",
                                code="21.01", revenue=Decimal("300"), cost=Decimal("180"))
        self.line4 = self._line(4, invoice_month=self.period.month, customer="客户甲",
                                category="纱线", code="11.01",
                                revenue=Decimal("400"), cost=Decimal("240"))

    def _line(self, row_no, *, out_month=None, invoice_month=None, customer="客户",
              category="纱线", code="11.01", revenue=Decimal("0"), cost=Decimal("0")):
        line = SalesLedgerLine.objects.create(
            period=self.period,
            batch=self.batch,
            row_no=row_no,
            out_month=out_month or self.period.month,
            invoice_status=SalesLedgerLine.InvoiceStatus.INVOICED,
            invoice_month=invoice_month,
            channel="国内销售",
            company="测试公司",
            order_no=f"SO-{row_no}",
            customer_code=f"C{row_no}",
            customer_name=customer,
            sales_category=category,
            material_code=code,
            material_name=f"产品{row_no}",
            uom="kg",
            quantity=Decimal("10"),
            terminal_amount=revenue,
            cost_amount=cost,
        )
        return line

    def _detail(self, url, qs=""):
        self.client.force_login(self.user)
        response = self.client.get(url + qs)
        self.assertEqual(response.status_code, 200)
        return response

    # ── 默认当月 ────────────────────────────────────────────────────────

    def test_defaults_to_current_month(self):
        url = reverse("profit_statement", args=[self.period.id])
        response = self._detail(url)
        detail = response.context["detail"]
        self.assertTrue(detail["default_applied"])
        self.assertEqual(detail["month"], "2026-07")
        self.assertEqual(detail["row_count"], 2)          # 只有 7 月两行
        self.assertEqual([r["seq"] for r in detail["rows"]], [3, 4])
        self.assertEqual(response.context["data"]["row_count"], 4)   # 台账全部 4 行
        html = response.content.decode("utf-8")
        self.assertIn("<h2>损益明细表</h2>", html)
        # 筛选控件都在，且都在同一条筛选行里（标题行不再有行数小字/标签）
        for name in ('name="dm"', 'name="ifrom"', 'name="ito"', 'name="cust"',
                     'name="dcat"', 'name="dcode"'):
            self.assertIn(name, html)

    def test_month_filter_lists_every_month(self):
        url = reverse("profit_statement", args=[self.period.id])
        options = self._detail(url).context["detail"]["month_options"]
        self.assertEqual([item["value"] for item in options], ["2026-06", "2026-07"])
        self.assertEqual([item["row_count"] for item in options], [2, 2])
        # 指定 6 月
        response = self._detail(url, "?df=1&dm=2026-06")
        self.assertEqual([r["seq"] for r in response.context["detail"]["rows"]], [1, 2])

    def test_clear_filter_shows_all_months(self):
        url = reverse("profit_statement", args=[self.period.id])
        response = self._detail(url, "?df=1&dm=&ifrom=&ito=&cust=&dcat=&dcode=")
        detail = response.context["detail"]
        self.assertEqual(detail["row_count"], 4)
        self.assertEqual(detail["month"], "")
        self.assertFalse(detail["has_filter"])

    # ── 开票期间区间 ────────────────────────────────────────────────────

    def test_invoice_month_range(self):
        url = reverse("profit_statement", args=[self.period.id])
        # 只看 6 月开票
        r = self._detail(url, "?df=1&dm=&ifrom=2026-06&ito=2026-06")
        self.assertEqual([x["seq"] for x in r.context["detail"]["rows"]], [1])
        # 6 月到 7 月
        r = self._detail(url, "?df=1&dm=&ifrom=2026-06&ito=2026-07")
        self.assertEqual([x["seq"] for x in r.context["detail"]["rows"]], [1, 2, 4])
        # 只填「从」= 该月及以后；只填「到」= 该月及以前
        r = self._detail(url, "?df=1&dm=&ifrom=2026-07")
        self.assertEqual([x["seq"] for x in r.context["detail"]["rows"]], [2, 4])
        r = self._detail(url, "?df=1&dm=&ito=2026-06")
        self.assertEqual([x["seq"] for x in r.context["detail"]["rows"]], [1])

    def test_invoice_range_includes_uninvoiced_and_swaps_bounds(self):
        url = reverse("profit_statement", args=[self.period.id])
        response = self._detail(url, "?df=1&dm=&ifrom=2026-07&ito=2026-06")
        detail = response.context["detail"]
        self.assertEqual((detail["invoice_from"], detail["invoice_to"]),
                         ("2026-06", "2026-07"))          # 起止填反自动对调
        self.assertEqual([x["seq"] for x in detail["rows"]], [1, 2, 4])
        # 「(未开票)」可被选中，且排在最后
        uninvoiced = "(\u672a\u5f00\u7968)"
        self.assertEqual(
            [item["value"] for item in detail["invoice_options"]][-1], uninvoiced
        )
        r = self._detail(url, "?df=1&dm=&ifrom=%28%E6%9C%AA%E5%BC%80%E7%A5%A8%29"
                              "&ito=%28%E6%9C%AA%E5%BC%80%E7%A5%A8%29")
        self.assertEqual([x["seq"] for x in r.context["detail"]["rows"]], [3])

    # ── 客户名称 / 销售大类 / 产品编码 ──────────────────────────────────

    def test_customer_name_filter(self):
        url = reverse("profit_statement", args=[self.period.id])
        r = self._detail(url, "?df=1&dm=&cust=%E5%AE%A2%E6%88%B7%E7%94%B2")
        self.assertEqual([x["seq"] for x in r.context["detail"]["rows"]], [1, 3, 4])
        # 空结果
        r = self._detail(url, "?df=1&dm=&cust=ZZZ")
        self.assertEqual(r.context["detail"]["row_count"], 0)

    def test_category_filter_multi(self):
        url = reverse("profit_statement", args=[self.period.id])
        r = self._detail(url, "?df=1&dm=&dcat=%E7%BA%B1%E7%BA%BF")
        self.assertEqual([x["seq"] for x in r.context["detail"]["rows"]], [1, 4])
        r = self._detail(url, "?df=1&dm=&dcat=%E7%BA%B1%E7%BA%BF&dcat=%E7%BA%A4%E7%BB%B4")
        self.assertEqual(r.context["detail"]["categories"], ["纱线", "纤维"])
        self.assertEqual([x["seq"] for x in r.context["detail"]["rows"]], [1, 3, 4])

    def test_product_code_filter(self):
        url = reverse("profit_statement", args=[self.period.id])
        r = self._detail(url, "?df=1&dm=&dcode=11.0")
        self.assertEqual([x["seq"] for x in r.context["detail"]["rows"]], [1, 2, 4])
        r = self._detail(url, "?df=1&dm=&dcode=21")
        self.assertEqual([x["seq"] for x in r.context["detail"]["rows"]], [3])

    def test_invalid_values_ignored(self):
        url = reverse("profit_statement", args=[self.period.id])
        r = self._detail(url, "?df=1&dm=2000-01&ifrom=xx&ito=yy&dcat=不存在")
        detail = r.context["detail"]
        self.assertEqual(detail["month"], "")
        self.assertEqual(detail["invoice_from"], "")
        self.assertEqual(detail["categories"], [])
        self.assertEqual(detail["row_count"], 4)

    # ── 组合 / 合计 / 与汇总互不影响 ────────────────────────────────────

    def test_combined_filters_and_totals_follow(self):
        url = reverse("profit_statement", args=[self.period.id])
        r = self._detail(url, "?df=1&dm=2026-07&dcat=%E7%BA%B1%E7%BA%BF&dcode=11")
        detail = r.context["detail"]
        self.assertEqual([x["seq"] for x in detail["rows"]], [4])
        self.assertEqual(detail["totals"]["revenue"], Decimal("400.00"))
        html = r.content.decode("utf-8")
        self.assertIn("合计（当前筛选）", html)
        # 合计随筛选变化：当月两行 vs 全部四行
        all_rows = self._detail(url, "?df=1&dm=")
        self.assertEqual(all_rows.context["detail"]["totals"]["revenue"], Decimal("1000.00"))

    def test_detail_filter_does_not_change_summary(self):
        url = reverse("profit_statement", args=[self.period.id])
        before = self._detail(url, "?basis=out&f=1&m=2026-07")
        after = self._detail(url, "?basis=out&f=1&m=2026-07&df=1&dm=2026-06&cust=X")
        self.assertEqual(
            before.context["summary"]["grand_total"]["revenue"],
            after.context["summary"]["grand_total"]["revenue"],
        )
        self.assertEqual(after.context["detail"]["row_count"], 0)   # 明细被筛掉了
        # 明细筛选不影响汇总表的期间行
        self.assertEqual(
            [row["month"] for row in after.context["summary"]["rows"]
             if row["level"] == "month"], ["2026-07"],
        )

    def test_summary_filter_does_not_change_detail_default(self):
        url = reverse("profit_statement", args=[self.period.id])
        r = self._detail(url, "?basis=out&f=1&m=2026-06")
        # 汇总看 6 月，明细仍按自己的默认（当月 7 月）
        self.assertEqual(r.context["selected_periods"], ["2026-06"])
        self.assertEqual(r.context["detail"]["month"], "2026-07")

    def test_pagination_keeps_filters(self):
        url = reverse("profit_statement", args=[self.period.id])
        r = self._detail(url, "?df=1&dm=2026-06&dcode=11")
        querystring = r.context["page"].querystring
        self.assertIn("dm=2026-06", querystring)
        self.assertIn("dcode=11", querystring)
        self.assertIn("df=1", querystring)
        # 清空筛选的链接把明细条件都留空（= 看全部月份）
        clear = r.context["detail_clear_query"]
        for key in ("dm=", "ifrom=", "ito=", "cust=", "dcat=", "dcode="):
            self.assertIn(key, clear)
        self.assertIn("df=1", clear)
        r2 = self._detail(url, "?" + clear)
        self.assertEqual(r2.context["detail"]["row_count"], 4)

    def test_detail_export_follows_filters(self):
        # 用户 2026-09-28：损益明细表加「导出Excel」，导出 = 当前筛选条件看到的那张明细表
        self.client.force_login(self.user)
        url = reverse("profit_detail_export", args=[self.period.id])

        # 不带条件：导出默认视图（当月 2026-07 的 2 行）+ 合计行
        response = self.client.get(url)
        self.assertEqual(response.status_code, 200)
        workbook = load_workbook(BytesIO(response.content))
        self.assertEqual(workbook.sheetnames, ["损益明细表"])
        sheet = workbook["损益明细表"]
        headers = [cell.value for cell in sheet[1]]
        self.assertEqual(headers[0], "序号")
        self.assertIn("终端销售收入", headers)
        rows = list(sheet.iter_rows(values_only=True))
        self.assertEqual(len(rows), 1 + 2 + 1)          # 表头 + 2 行明细 + 合计
        self.assertEqual(rows[-1][0], "合计")

        # 带条件：出库期间 6 月 + 客户甲 → 只剩 1 行
        response = self.client.get(url + "?df=1&dm=2026-06&cust=%E5%AE%A2%E6%88%B7%E7%94%B2")
        workbook = load_workbook(BytesIO(response.content))
        sheet = workbook["损益明细表"]
        headers = [c.value for c in sheet[1]]
        rows = list(sheet.iter_rows(values_only=True))
        self.assertEqual(len(rows), 1 + 1 + 1)
        self.assertEqual(rows[1][headers.index("客户名称")], "客户甲")
        self.assertEqual(Decimal(str(rows[-1][headers.index("终端销售收入")])),
                         Decimal("100.00"))

        # 放开全部月份：4 行
        response = self.client.get(url + "?df=1")
        workbook = load_workbook(BytesIO(response.content))
        rows = list(workbook["损益明细表"].iter_rows(values_only=True))
        self.assertEqual(len(rows), 1 + 4 + 1)

    def test_detail_filter_row_is_single_line_with_export_and_print(self):
        # 筛选字段 + 确定/清空筛选/导出Excel/打印本表 都在同一条筛选行里（底部对齐）
        response = self._detail(reverse("profit_statement", args=[self.period.id]))
        html = response.content.decode("utf-8")
        self.assertIn('class="filter-grid compact-filters detail-filters"', html)
        section = html[html.index("<h2>损益明细表</h2>"):]
        section = section[: section.index("</section>")]
        form = section.split('class="filter-grid compact-filters detail-filters"', 1)[1].split("</form>", 1)[0]
        for name in ('name="dm"', 'name="ifrom"', 'name="ito"', 'name="cust"',
                     'name="dcat"', 'name="dcode"'):
            self.assertIn(name, form, name)
        self.assertIn('class="filter-actions"', form)
        self.assertIn("导出Excel", form)
        self.assertIn(reverse("profit_detail_export", args=[self.period.id]), form)
        # 「打印本表」的注入位与筛选字段同一行
        self.assertIn('class="print-here"', form)
        # 导出链接带上当前明细筛选
        self.assertIn("dm=2026-07", form)

    def test_month_option_hides_row_count(self):
        # 用户 2026-09-28：「出库期间」下拉里的「（N 行）」数字不再显示
        response = self._detail(reverse("profit_statement", args=[self.period.id]))
        html = response.content.decode("utf-8")
        section = html[html.index('class="filter-grid compact-filters detail-filters"'):]
        form = section.split("</form>", 1)[0]
        options = re.findall(r'<option value="([^"]+)"[^>]*>([^<]*)</option>',
                             form.split('name="dm"', 1)[1].split("</select>", 1)[0])
        self.assertTrue(options)
        for value, label in options:
            self.assertNotIn("行）", label, label)
            self.assertNotIn("行)", label, label)
        # 下拉里仍能列出每个出库月份，且「全部」选项还在
        labels = [label.strip() for _v, label in options]
        self.assertIn("2026-07", labels)
        self.assertIn('<option value="">全部</option>', form)

    def test_detail_filter_row_styles_single_line_and_shorter_boxes(self):
        # 样式：筛选行是 flex 一行（按钮组整体换行），方框比通用筛选栏略窄略矮
        css = (settings.BASE_DIR / "static" / "settlement" / "app.css").read_text(
            encoding="utf-8"
        )
        block = css.split("form.filter-grid.compact-filters {", 1)[1].split("}", 1)[0]
        self.assertIn("display: flex", block)
        self.assertIn("flex-wrap: wrap", block)
        self.assertIn("align-items: flex-end", block)
        self.assertIn(
            "form.filter-grid.compact-filters .filter-field {\n  flex: 1 1 110px;", css
        )
        self.assertIn("max-width: 172px", css)
        self.assertIn("min-height: 34px", css)

    def test_saving_expenses_keeps_detail_filters(self):
        url = reverse("profit_statement", args=[self.period.id])
        self.client.force_login(self.user)
        response = self.client.post(
            url + "?basis=out&df=1&dm=2026-06&cust=%E5%AE%A2%E6%88%B7%E7%94%B2",
            {"exp_2026-07_admin_expense": "0"},
        )
        location = response["Location"]
        self.assertIn("dm=2026-06", location)
        self.assertIn("cust=", location)
        self.assertIn("df=1", location)

    # ── 服务层函数 ──────────────────────────────────────────────────────

    def test_service_helpers(self):
        data = build_profit_statement(self.period)
        rows = data["rows"]
        options = detail_month_options(rows)
        self.assertEqual([item["value"] for item in options], ["2026-06", "2026-07"])
        self.assertEqual(options[0]["row_count"], 2)

        kept = filter_detail_rows(rows, month="2026-07")
        self.assertEqual(len(kept), 2)
        kept = filter_detail_rows(rows, customer="客户甲", categories=["纱线"])
        self.assertEqual([r["seq"] for r in kept], [1, 4])
        kept = filter_detail_rows(rows, code="21")
        self.assertEqual([r["seq"] for r in kept], [3])
        kept = filter_detail_rows(rows, invoice_from="2026-06", invoice_to="2026-06")
        self.assertEqual([r["seq"] for r in kept], [1])
        # 空条件 = 不筛
        self.assertEqual(len(filter_detail_rows(rows)), 4)
        # 合计
        totals = row_totals(kept)
        self.assertEqual(totals["revenue"], Decimal("100.00"))


class ExpenseSaveSafetyTests(SettlementFixtureMixin, TestCase):
    """保存费用：只更新「提交里确实出现」的字段，绝不把未提交的录入值清成 0。

    背景（2026-09-15）：一次只带一个字段的提交把 7 月已录数据清成了 0，
    因此加此保护：过期页面/半截提交不再破坏已有数据。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="expense-safety", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.batch = self.create_batch(self.period, ImportBatch.Kind.SALES_LEDGER, "safe")
        for row_no in (1, 2):
            SalesLedgerLine.objects.create(
                period=self.period, batch=self.batch, row_no=row_no,
                out_month=self.period.month,
                invoice_status=SalesLedgerLine.InvoiceStatus.INVOICED,
                invoice_month=self.period.month, channel="国内销售",
                order_no=f"SO-{row_no}", customer_name="客户", sales_category="纱线",
                material_code="11.01", material_name="产品", uom="kg",
                quantity=Decimal("10"), terminal_amount=Decimal("1000"),
                cost_amount=Decimal("600"),
            )
        self.url = reverse("profit_statement", args=[self.period.id])
        # 7 月已录好一套费用
        self.expense = ProfitStatementExpense.objects.create(
            period=self.period, month=self.period.month,
            admin_expense=Decimal("2098509.12"), selling_expense=Decimal("265454.86"),
            finance_expense=Decimal("169536.42"), non_operating_expense=Decimal("992.84"),
        )

    def _fresh(self):
        return ProfitStatementExpense.objects.get(
            period=self.period, month=self.period.month
        )

    def test_partial_submit_does_not_clear_other_fields(self):
        # 只提交一个字段（模拟半截提交）：其它字段必须保持原值
        self.client.force_login(self.user)
        self.client.post(self.url, {"exp_2026-07_income_tax": "123"})
        saved = self._fresh()
        self.assertEqual(saved.income_tax, Decimal("123.00"))
        self.assertEqual(saved.admin_expense, Decimal("2098509.12"))
        self.assertEqual(saved.selling_expense, Decimal("265454.86"))
        self.assertEqual(saved.finance_expense, Decimal("169536.42"))
        self.assertEqual(saved.non_operating_expense, Decimal("992.84"))

    def test_empty_submit_changes_nothing(self):
        # 一个字段都没提交：直接不动数据
        self.client.force_login(self.user)
        response = self.client.post(self.url, {})
        self.assertEqual(response.status_code, 302)
        saved = self._fresh()
        self.assertEqual(saved.admin_expense, Decimal("2098509.12"))
        self.assertEqual(saved.selling_expense, Decimal("265454.86"))

    def test_explicit_empty_value_still_clears(self):
        # 字段出现但留空 = 用户主动清空 → 仍按 0 处理（保留原语义）
        self.client.force_login(self.user)
        self.client.post(self.url, {"exp_2026-07_admin_expense": ""})
        self.assertEqual(self._fresh().admin_expense, Decimal("0.00"))
        self.assertEqual(self._fresh().selling_expense, Decimal("265454.86"))

    def test_full_form_submit_still_saves_everything(self):
        # 完整表单提交（页面正常路径）照旧全部保存
        self.client.force_login(self.user)
        payload = {}
        for field, _label in ALLOCATION_FIELDS:
            payload[f"exp_2026-07_{field}"] = "1,000.50"
        self.client.post(self.url, payload)
        saved = self._fresh()
        for field, _label in ALLOCATION_FIELDS:
            self.assertEqual(getattr(saved, field), Decimal("1000.50"), field)

    def test_expense_table_row_height_and_grid(self):
        # 费用表：行高放宽（不跟随 25 磅）、网格线加深
        css = (settings.BASE_DIR / "static" / "settlement" / "app.css").read_text(
            encoding="utf-8"
        )
        blocks = re.findall(
            r"\.expense-entry-table th, \.expense-entry-table td \{([^}]*)\}", css
        )
        self.assertTrue(blocks, "缺少费用表单元格样式")
        body = "\n".join(blocks)      # 该选择器可能有多条规则，合起来校验
        self.assertIn("min-height: 46px", body)      # 行高适中（一屏放得下）
        self.assertIn("height: auto", body)          # 不受全局 25pt 限制
        self.assertIn("border-color: #cfd8de", body)  # 网格线加深
        # 输入框：不再单独画小方框（无边框、透明底），宽度撑满格子
        inputs = re.findall(
            r"\.expense-entry-table input\.expense-input \{([^}]*)\}", css
        )
        merged = "\n".join(inputs)
        self.assertIn("border: 0", merged)
        self.assertIn("background: transparent", merged)


class ExpenseTableLayoutTests(SettlementFixtureMixin, TestCase):
    """按月费用录入表：整张表一屏放下；网格线已分格，输入框不再单独画小方框。"""

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="expense-layout", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.batch = self.create_batch(self.period, ImportBatch.Kind.SALES_LEDGER, "lay")
        for row_no in (1, 2):
            SalesLedgerLine.objects.create(
                period=self.period, batch=self.batch, row_no=row_no,
                out_month=self.period.month,
                invoice_status=SalesLedgerLine.InvoiceStatus.INVOICED,
                invoice_month=self.period.month, channel="国内销售",
                order_no=f"SO-{row_no}", customer_name="客户", sales_category="纱线",
                material_code="11.01", material_name="产品", uom="kg",
                quantity=Decimal("10"), terminal_amount=Decimal("1000"),
                cost_amount=Decimal("600"),
            )
        self.css = (settings.BASE_DIR / "static" / "settlement" / "app.css").read_text(
            encoding="utf-8"
        )

    def test_table_fits_container_no_horizontal_scroll(self):
        # 用 table-layout: fixed + 100% 宽，列按容器等分，所以不会超出容器
        block = re.search(
            r"\.table-scroll > \.expense-entry-table \{([^}]*)\}", self.css
        )
        self.assertIsNotNone(block, "缺少费用表宽度规则")
        body = block.group(1)
        self.assertIn("width: 100%", body)
        self.assertIn("min-width: 0", body)      # 覆盖全局 min-width:100% 的换行撑宽
        self.assertIn("table-layout: fixed", body)

    def test_number_columns_are_narrow(self):
        # 数字列收窄，避免空荡荡的宽格子（原来每列 ~196px）
        block = re.search(
            r"\.expense-entry-table th \+ th,\s*\.expense-entry-table td \+ td \{([^}]*)\}",
            self.css,
        )
        self.assertIsNotNone(block, "缺少数字列宽规则")
        width = int(re.search(r"width: (\d+)px", block.group(1)).group(1))
        self.assertLessEqual(width, 150)
        self.assertGreaterEqual(width, 110)      # 仍要放得下 2,098,509.12

    def test_input_has_no_separate_box(self):
        # 输入框无边框、透明底：靠单元格网格线分格，不再单独画方框
        blocks = re.findall(
            r"\.expense-entry-table input\.expense-input \{([^}]*)\}", self.css
        )
        self.assertTrue(blocks, "缺少费用输入框样式")
        body = "\n".join(blocks)   # 该选择器可能有多条规则，合并校验
        self.assertIn("border: 0", body)
        self.assertIn("background: transparent", body)
        self.assertIn("width: 100%", body)
        # 聚焦时才给个提示，仍然不画方框（用内阴影）
        focus = re.search(
            r"\.expense-entry-table input\.expense-input:focus \{([^}]*)\}", self.css
        )
        self.assertIsNotNone(focus)
        self.assertIn("box-shadow", focus.group(1))
        self.assertIn("outline: none", focus.group(1))

    def test_row_height_reasonable_and_no_vertical_scroll(self):
        # 行高适中（不跟数据列表压成 25 磅，也不回到 56px 那么高）
        block = re.search(
            r"\.expense-entry-table th, \.expense-entry-table td \{([^}]*)\}", self.css
        )
        self.assertIn("min-height: 46px", block.group(1))
        self.assertIn("height: auto", block.group(1))

    def test_timestamp_in_period_cell_is_short(self):
        # 期间列里的录入时间用「月-日 时:分」，避免该格被省略号截断
        tpl = (settings.BASE_DIR / "templates" / "settlement"
               / "profit_statement.html").read_text(encoding="utf-8")
        self.assertIn('item.updated_at|date:"m-d H:i"', tpl)
        self.assertNotIn('item.updated_at|date:"Y-m-d H:i"', tpl)


class PrintSupportTests(SettlementFixtureMixin, TestCase):
    """打印功能：页面上有打印按钮；打印时隐藏外壳、展开滚动框、内容不截断、表头跨页重复。"""

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="print-tester", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.base_tpl = (settings.BASE_DIR / "templates" / "base.html").read_text(
            encoding="utf-8"
        )
        self.css = (settings.BASE_DIR / "static" / "settlement" / "app.css").read_text(
            encoding="utf-8"
        )

    def test_section_print_buttons_are_added_next_to_each_table(self):
        # 打印按钮放在「想打印的那一块」旁边：自动为每个含表格的板块注入「打印本表」
        self.assertIn("window.printSection", self.base_tpl)
        self.assertIn("addSectionPrintButtons", self.base_tpl)
        self.assertIn("section-print", self.base_tpl)
        # 没有标题行的板块也会补一个标题行，按钮才有地方放
        self.assertIn("section-heading", self.base_tpl)
        # 按钮本身 no-print，打印时不会被印出来
        self.assertIn("no-print", self.base_tpl)
        self.assertIn(".section-print { white-space: nowrap; }", self.css)

    def test_topbar_has_no_global_print_button(self):
        # 顶部不再放整页打印按钮（用户要求放在具体板块旁边）
        self.assertNotIn('class="link-button no-print" type="button" onclick="window.print()"',
                         self.base_tpl)

    def test_print_heading_still_present(self):
        # 打印抬头在页面上不显示、打印时才出现（含打印时间）
        self.assertIn('class="print-heading"', self.base_tpl)
        self.assertIn(".print-heading { display: none; }", self.css)

    def test_print_heading_filled_by_script(self):
        # 抬头里的标题/时间由脚本填（含打印前 beforeprint 兜底）
        self.assertIn("fillPrintHeading", self.base_tpl)
        self.assertIn("beforeprint", self.base_tpl)

    def test_print_css_hides_shell_and_expands_scroll_containers(self):
        block = re.search(r"@media print \{(.*)\n\}\n\Z", self.css, re.S)
        self.assertIsNotNone(block, "缺少 @media print 样式")
        css = block.group(1)
        # 隐藏导航/筛选/分页/按钮
        for sel in (".topbar", ".period-nav", ".pagination", ".no-print", "form.filter-grid"):
            self.assertIn(sel, css)
        # 滚动容器必须展开，否则只会打印框里看得见的一屏
        for sel in (".table-wrap", ".table-scroll", ".profit-table-wrap", ".summary-table-wrap"):
            self.assertIn(sel, css)
        self.assertIn("overflow: visible !important", css)
        # 单元格不再用省略号截断（否则打印会丢内容）
        self.assertIn("text-overflow: clip !important", css)
        self.assertIn("white-space: normal !important", css)
        # 表头跨页重复
        self.assertIn("display: table-header-group", css)
        # A4 横向
        self.assertIn("size: A4 landscape", css)
        # 单块打印：点某块的按钮时只打印那一块
        self.assertIn("body.print-focus", css)
        self.assertIn(".is-print-target", css)

    def test_profit_statement_page_has_print_button(self):
        self.client.force_login(self.user)
        sales_line = SalesLedgerLine.objects.create(
            period=self.period,
            batch=self.create_batch(self.period, ImportBatch.Kind.SALES_LEDGER, "pr"),
            row_no=1, out_month=self.period.month,
            invoice_status=SalesLedgerLine.InvoiceStatus.INVOICED,
            invoice_month=self.period.month, channel="国内销售",
            order_no="SO-1", customer_name="客户", sales_category="纱线",
            material_code="11.01", material_name="产品", uom="kg",
            quantity=Decimal("1"), terminal_amount=Decimal("100"),
            cost_amount=Decimal("60"),
        )
        self.assertTrue(sales_line.pk)
        html = self.client.get(
            reverse("profit_statement", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertIn("window.print()", html)
        self.assertIn('class="print-heading"', html)

    def test_every_period_page_gets_the_print_button(self):
        # 打印按钮在公共模板里，所以带期间导航的页面都有
        self.client.force_login(self.user)
        for name in ("inventory_ledger", "sales_uninvoiced",
                     "purchase_contracts", "reconciliation"):
            html = self.client.get(
                reverse(name, args=[self.period.id])
            ).content.decode("utf-8")
            self.assertIn("window.print()", html, name)
            self.assertIn('class="print-heading"', html, name)


class SalesVoucherAndStatisticsTests(SettlementFixtureMixin, TestCase):
    """销售结转成本凭证（独立页面）+ 发出商品及销售成本统计表（仍在月末未开票页）。"""

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="voucher-tester", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.batch = self.create_batch(self.period, ImportBatch.Kind.SALES_LEDGER, "vch")
        self.may, self.june = date(2026, 5, 1), date(2026, 6, 1)

    def _line(self, row_no, *, out_month, invoice_month=None, carry_type,
              status="invoiced", cost="100.00", quantity="10", category="纱线"):
        return SalesLedgerLine.objects.create(
            period=self.period, batch=self.batch, row_no=row_no,
            out_month=out_month,
            invoice_status=status,
            invoice_month=invoice_month,
            channel="国内销售", order_no=f"SO-{row_no}",
            customer_code="C1", customer_name="客户甲", sales_category=category,
            material_code="11.01", material_name="产品", uom="kg",
            quantity=Decimal(quantity), terminal_amount=Decimal("200.00"),
            cost_amount=Decimal(cost), carry_type=carry_type,
        )

    # ── 销售结转成本凭证（按销售大类汇总；筛选在表内单元格里） ──────────

    def test_voucher_summary_groups_by_category(self):
        # 统计三类结转；每个结转类型都列全 5 个销售大类（无数据显示 0）
        self._line(1, out_month=self.period.month, invoice_month=self.period.month,
                   carry_type="当月发货当月开票结转成本", cost="100.00", quantity="10",
                   category="纱线")
        self._line(2, out_month=self.period.month, invoice_month=self.period.month,
                   carry_type="当月发货当月开票结转成本", cost="50.00", quantity="5",
                   category="纱线")
        self._line(3, out_month=self.period.month, invoice_month=self.period.month,
                   carry_type="当月发货当月开票结转成本", cost="300.00", quantity="20",
                   category="面料")
        self._line(4, out_month=self.june, invoice_month=self.period.month,
                   carry_type="前期发货本期开票结转成本", cost="700.00", quantity="40",
                   category="纤维")
        self._line(5, out_month=self.period.month, carry_type="当月发货未开票挂发出商品",
                   status="uninvoiced", cost="900.00", quantity="30", category="面料")
        # 免费样品不在凭证统计范围内
        self._line(6, out_month=self.period.month, carry_type="免费样品赠送计入销售费用",
                   status="sample_free", cost="888.00")

        data = build_sales_voucher_summary(self.period)
        cats = [row for row in data["rows"] if row["level"] == "category"]
        # 3 个结转类型 × 5 个销售大类 = 15 行
        self.assertEqual(len(cats), 15)
        got = {(row["carry_type"], row["category"]): str(row["amount"]) for row in cats}
        self.assertEqual(got[("当月发货当月开票结转成本", "纱线")], "150.00")
        self.assertEqual(got[("当月发货当月开票结转成本", "面料")], "300.00")
        self.assertEqual(got[("前期发货本期开票结转成本", "纤维")], "700.00")
        self.assertEqual(got[("当月发货未开票挂发出商品", "面料")], "900.00")
        # 没有数据的大类也要出现（0）
        for carry in ("当月发货当月开票结转成本", "前期发货本期开票结转成本",
                      "当月发货未开票挂发出商品"):
            for category in ("其他成品", "原料"):
                self.assertEqual(got[(carry, category)], "0.00")
        # 每个结转类型的销售大类顺序固定为 5 个
        for carry in ("当月发货当月开票结转成本", "前期发货本期开票结转成本",
                      "当月发货未开票挂发出商品"):
            order = [row["category"] for row in cats if row["carry_type"] == carry]
            self.assertEqual(order, ["纱线", "面料", "纤维", "其他成品", "原料"], carry)
        # 会计科目：两类已开票 → 主营业务成本；未开票挂发出商品 → 发出商品
        accounts = {row["carry_type"]: row["account"] for row in cats}
        self.assertEqual(accounts["当月发货当月开票结转成本"], "主营业务成本")
        self.assertEqual(accounts["前期发货本期开票结转成本"], "主营业务成本")
        self.assertEqual(accounts["当月发货未开票挂发出商品"], "发出商品")
        # 小计（3 个）与合计
        subtotals = {row["carry_type"]: str(row["amount"])
                     for row in data["rows"] if row["level"] == "subtotal"}
        self.assertEqual(subtotals, {
            "当月发货当月开票结转成本": "450.00",
            "前期发货本期开票结转成本": "700.00",
            "当月发货未开票挂发出商品": "900.00",
        })
        total = [row for row in data["rows"] if row["level"] == "total"][0]
        self.assertEqual(str(total["amount"]), "2050.00")
        # 合计行没有单一会计科目 → 留空
        self.assertEqual(total["account"], "")
        # 列头已去掉「抬头销售凭证」
        self.assertNotIn("抬头销售凭证", data["headers"])
        self.assertEqual(data["headers"][:3], ("会计科目", "结转类型", "销售大类"))

    def test_voucher_summary_defaults_per_block(self):
        # 三块默认值：前两块 = 出库期间空白 + 已开发票 + 当月开票期间；第三块 = 当月出库 + 空白/空白
        data = build_sales_voucher_summary(self.period)
        block_values = {b["index"]: b for b in data["blocks"]}
        self.assertEqual(block_values[0]["out_value"], "")
        self.assertEqual(block_values[0]["status_value"], "invoiced")
        self.assertEqual(block_values[0]["invoice_value"], "2026-07")
        self.assertEqual(block_values[1]["out_value"], "")
        self.assertEqual(block_values[1]["status_value"], "invoiced")
        self.assertEqual(block_values[1]["invoice_value"], "2026-07")
        self.assertEqual(block_values[2]["out_value"], "2026-07")
        self.assertEqual(block_values[2]["status_value"], "")
        self.assertEqual(block_values[2]["invoice_value"], "")
        # 三块的参数名互不相同（各自独立）
        names = {b["params"]["out"] for b in data["blocks"]}
        self.assertEqual(names, {"vout0", "vout1", "vout2"})

    def test_voucher_summary_blocks_are_independent(self):
        # 行先按「记账月份」归属：当月发货当月开票按出库月、前期发货本期开票按开票月。
        # 所以「前期发货本期开票」可以有不同出库月、但都落在 7 月记账 → 用它验证块内筛选。
        self._line(1, out_month=self.period.month, invoice_month=self.period.month,
                   carry_type="当月发货当月开票结转成本", cost="100.00", category="纱线")
        self._line(2, out_month=self.june, invoice_month=self.period.month,
                   carry_type="前期发货本期开票结转成本", cost="200.00", category="面料")
        self._line(3, out_month=self.may, invoice_month=self.period.month,
                   carry_type="前期发货本期开票结转成本", cost="300.00", category="纤维")
        self._line(4, out_month=self.period.month, carry_type="当月发货未开票挂发出商品",
                   status="uninvoiced", cost="800.00", category="原料")

        def subtotal(data, carry):
            rows = [r for r in data["rows"]
                    if r["level"] == "subtotal" and r["carry_type"] == carry]
            return str(rows[0]["amount"])

        BLOCK0 = "当月发货当月开票结转成本"
        BLOCK1 = "前期发货本期开票结转成本"
        BLOCK2 = "当月发货未开票挂发出商品"

        base = build_sales_voucher_summary(self.period)
        self.assertEqual(subtotal(base, BLOCK0), "100.00")
        self.assertEqual(subtotal(base, BLOCK1), "500.00")   # 6 月 200 + 5 月 300
        self.assertEqual(subtotal(base, BLOCK2), "800.00")

        # 只把块1 的出库期间改成 2026-06 → 块1 变 200；块0 / 块2 完全不变
        changed = build_sales_voucher_summary(self.period, block_filters={
            BLOCK1: {"out": "2026-06"},
        })
        self.assertEqual(subtotal(changed, BLOCK1), "200.00")
        self.assertEqual(subtotal(changed, BLOCK0), "100.00")
        self.assertEqual(subtotal(changed, BLOCK2), "800.00")

        # 只把块0 的开票状态改成未开发票 → 块0 变 0；块1 / 块2 不变
        changed0 = build_sales_voucher_summary(self.period, block_filters={
            BLOCK0: {"status": "uninvoiced"},
        })
        self.assertEqual(subtotal(changed0, BLOCK0), "0.00")
        self.assertEqual(subtotal(changed0, BLOCK1), "500.00")
        self.assertEqual(subtotal(changed0, BLOCK2), "800.00")

        # 只把块2 的三个筛选都放开 → 块2 仍 800（该块按出库月记账，结果不变）；块0 / 块1 不变
        changed2 = build_sales_voucher_summary(self.period, block_filters={
            BLOCK2: {"out": "", "status": "", "invoice": ""},
        })
        self.assertEqual(subtotal(changed2, BLOCK2), "800.00")
        self.assertEqual(subtotal(changed2, BLOCK0), "100.00")
        self.assertEqual(subtotal(changed2, BLOCK1), "500.00")

    def test_voucher_summary_blank_means_no_restriction(self):
        self._line(1, out_month=self.june, invoice_month=self.period.month,
                   carry_type="前期发货本期开票结转成本", cost="200.00", category="面料")
        self._line(2, out_month=self.may, invoice_month=self.period.month,
                   carry_type="前期发货本期开票结转成本", cost="300.00", category="纤维")
        BLOCK1 = "前期发货本期开票结转成本"

        def subtotal_of(data):
            rows = [r for r in data["rows"]
                    if r["level"] == "subtotal" and r["carry_type"] == BLOCK1]
            return str(rows[0]["amount"])

        # 默认（出库期间空白）→ 两行都算进来
        self.assertEqual(subtotal_of(build_sales_voucher_summary(self.period)), "500.00")
        # 出库期间选具体月份 → 只算该月
        self.assertEqual(
            subtotal_of(build_sales_voucher_summary(self.period, block_filters={
                BLOCK1: {"out": "2026-06"},
            })), "200.00")
        # 出库期间选「（空白）」→ 恢复不限制
        self.assertEqual(
            subtotal_of(build_sales_voucher_summary(self.period, block_filters={
                BLOCK1: {"out": ""},
            })), "500.00")
        # 非法取值按「不限制」处理
        self.assertEqual(
            subtotal_of(build_sales_voucher_summary(self.period, block_filters={
                BLOCK1: {"out": "1999-01", "status": "zzz", "invoice": "1999-01"},
            })), "500.00")

    def test_voucher_page_renders_entries_and_block_filters(self):
        """凭证页＝凭证分录格式：借贷列、三块取数条件、附件里的三块独立筛选。"""
        self._line(1, out_month=self.period.month, invoice_month=self.period.month,
                   carry_type="当月发货当月开票结转成本", cost="100.00", category="纱线")
        self.client.force_login(self.user)
        # 凭证已拆成独立页面
        url = reverse("sales_cost_voucher", args=[self.period.id])
        html = self.client.get(url).content.decode("utf-8")
        self.assertIn("<h1>销售结转成本凭证</h1>", html)
        self.assertIn("<h2>凭证分录（不含税）</h2>", html)
        self.assertIn("<h2>附件：结转类型 × 销售大类（含表内筛选）</h2>", html)
        # 表头＝凭证分录口径
        idx = html.index("<h2>凭证分录（不含税）</h2>")
        section = html[idx:html.index("</table>", idx)]
        heads = [re.sub(r"<[^>]+>", "", h).strip()
                 for h in re.findall(r"<th[^>]*>(.*?)</th>", section, re.S)]
        self.assertEqual(heads, ["摘要", "科目", "数量", "借方", "贷方"])
        # 附件里三块的下拉参数名互不相同（保留原有独立筛选）
        for name in ("vout0", "vst0", "vinv0", "vout1", "vst1", "vinv1",
                     "vout2", "vst2", "vinv2"):
            self.assertIn('name="%s"' % name, html)
        # 默认选中项：块0 已开发票 + 当月
        self.assertIn('value="invoiced" selected', html)
        # 带出借贷科目
        self.assertIn("库存商品", html)

    def test_voucher_entries_accounts_and_balance(self):
        """三类结转的借贷科目与金额：借＝贷，转类型决定贷方是库存商品还是发出商品。"""
        self._line(1, out_month=self.period.month, invoice_month=self.period.month,
                   carry_type="当月发货当月开票结转成本", cost="100.00", category="纱线")
        self._line(2, out_month=self.june, invoice_month=self.period.month,
                   carry_type="前期发货本期开票结转成本", cost="200.00", category="面料")
        self._line(3, out_month=self.period.month,
                   carry_type="当月发货未开票挂发出商品", status="uninvoiced",
                   cost="300.00", category="纤维")
        # 同一块再来一个销售大类，用来区分「按大类展开」与「科目汇总」
        self._line(4, out_month=self.period.month, invoice_month=self.period.month,
                   carry_type="当月发货当月开票结转成本", cost="50.00", category="面料")
        data = build_sales_voucher_entries(self.period)
        self.assertTrue(data["totals"]["balanced"])
        self.assertEqual(data["totals"]["debit"], Decimal("650.00"))
        self.assertEqual(data["totals"]["credit"], Decimal("650.00"))
        blocks = {b["carry_type"]: b for b in data["blocks"]}
        self.assertEqual(blocks["当月发货当月开票结转成本"]["debit_account"], "主营业务成本")
        self.assertEqual(blocks["当月发货当月开票结转成本"]["credit_account"], "库存商品")
        self.assertEqual(blocks["前期发货本期开票结转成本"]["debit_account"], "主营业务成本")
        self.assertEqual(blocks["前期发货本期开票结转成本"]["credit_account"], "发出商品")
        self.assertEqual(blocks["当月发货未开票挂发出商品"]["debit_account"], "发出商品")
        self.assertEqual(blocks["当月发货未开票挂发出商品"]["credit_account"], "库存商品")
        # 借方只到销售渠道（销售大类体现在贷方与附件里）
        # 借方「主营业务成本」只按销售渠道分国内/国外（不带销售大类），每块合并一行
        categories = [
            row["account"]
            for row in blocks["当月发货当月开票结转成本"]["entries"]
            if row["level"] == "line" and row["debit"] > 0
        ]
        self.assertEqual(categories, ["主营业务成本—国内销售"])
        # 附件表并列显示借贷科目
        self.assertEqual(
            list(data["summary_headers"])[:2], ["会计科目", "贷方科目"]
        )
        for row in data["summary_rows"]:
            if row["level"] == "category" and row["block_index"] >= 0:
                self.assertTrue(row["credit_account"])
        # 科目汇总模式：每块借贷各一行
        single = build_sales_voucher_entries(self.period, group="single")
        self.assertTrue(single["totals"]["balanced"])
        # 科目汇总模式：每块借贷各一行，比按大类展开少行
        self.assertLess(
            single["totals"]["line_count"], data["totals"]["line_count"]
        )
        self.assertEqual(single["totals"]["debit"], data["totals"]["debit"])
        # 贷方也按销售大类展开
        self.assertEqual(
            {
                row["account"]
                for row in blocks["当月发货当月开票结转成本"]["entries"]
                if row["level"] == "line" and row["credit"] > 0
            },
            {"库存商品—纱线", "库存商品—面料"},
        )
        self.assertEqual(
            {
                row["account"]
                for row in blocks["前期发货本期开票结转成本"]["entries"]
                if row["level"] == "line" and row["credit"] > 0
            },
            {"发出商品—面料"},
        )
        self.assertEqual(
            {
                row["account"]
                for row in blocks["当月发货未开票挂发出商品"]["entries"]
                if row["level"] == "line" and row["credit"] > 0
            },
            {"库存商品—纤维"},
        )
        # 借方是发出商品，不按渠道拆
        self.assertEqual(
            {
                row["account"]
                for row in blocks["当月发货未开票挂发出商品"]["entries"]
                if row["level"] == "line" and row["debit"] > 0
            },
            {"发出商品—纤维"},
        )
        # 借贷两侧逐大类金额一致
        block1 = blocks["当月发货当月开票结转成本"]
        # 借贷两侧合计一致（借方按渠道、贷方按销售大类）
        debit_sum = sum(
            (row["debit"] for row in block1["entries"] if row["level"] == "line"),
            Decimal("0"),
        )
        credit_sum = sum(
            (row["credit"] for row in block1["entries"] if row["level"] == "line"),
            Decimal("0"),
        )
        self.assertEqual(debit_sum, credit_sum)


    def test_sales_voucher_slim_format(self):
        """新格式：摘要不带销售大类、无辅助核算/来源行数、数量在借方前、凭证期间可筛选。"""
        self._line(1, out_month=self.period.month, invoice_month=self.period.month,
                   carry_type="当月发货当月开票结转成本", cost="100.00", category="纱线")
        data = build_sales_voucher_entries(self.period)
        self.assertEqual(list(data["headers"]), ["摘要", "科目", "数量", "借方", "贷方"])
        for block in data["blocks"]:
            for row in block["entries"]:
                # 摘要只有结转类型，不带销售大类
                self.assertNotIn("纱线", row["summary"])
                if row["level"] == "line" and row["account"]:
                    self.assertIn("—", row["account"])
        # 凭证中心卡片：无汇总字段、无借贷 pill、凭证期间是筛选
        center = build_voucher_center(self.period)
        card = {item["key"]: item for item in center["cards"]}["sales"]
        self.assertEqual(card["metrics"], [])
        self.assertIsNone(card["balance"])
        self.assertEqual(
            card["entry_headers"], ["摘要", "科目", "数量", "借方", "贷方"]
        )
        self.assertEqual(card["month_filter"]["param"], "sales_month")
        self.assertTrue(card["month_filter"]["options"])
        # 凭证中心页面渲染
        self.client.force_login(self.user)
        html = self.client.get(
            reverse("voucher_center", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertIn('name="sales_month"', html)
        self.assertIn("凭证期间", html)
        self.assertNotIn("参与结转台账行数", html)
        # 切期间后卡片只出该期间的数据
        june = {item["key"]: item
                for item in build_voucher_center(
                    self.period, sales_month="2026-06")["cards"]}["sales"]
        self.assertEqual(june["month"], "2026-06")

    def test_voucher_keeps_raw_sales_category(self):
        """销售大类保留台账原值：带括号的写法不能被归一成「其他成品」。"""
        self._line(1, out_month=self.period.month, invoice_month=self.period.month,
                   carry_type="当月发货当月开票结转成本", cost="100.00",
                   category="其他成品（手套、服装、床垫等）")
        self._line(2, out_month=self.period.month, invoice_month=self.period.month,
                   carry_type="当月发货当月开票结转成本", cost="50.00",
                   category="其他成品（毯子）")
        data = build_sales_voucher_entries(self.period)
        # 借方只到渠道；销售大类体现在贷方（去掉结尾括号）
        debit_accounts = {
            row["account"]
            for block in data["blocks"]
            for row in block["entries"]
            if row["level"] == "line" and row["debit"] > 0
        }
        self.assertEqual(debit_accounts, {"主营业务成本—国内销售"})
        credit_accounts = {
            row["account"]
            for block in data["blocks"]
            for row in block["entries"]
            if row["level"] == "line" and row["credit"] > 0
        }
        self.assertEqual(credit_accounts, {"库存商品—其他成品"})
        # 两个原值各自成行（分组没有被合并），只是显示名都去掉了括号
        raw_categories = {
            row["category"]
            for row in data["summary_rows"]
            if row["level"] == "category"
        }
        self.assertEqual(
            raw_categories, {"其他成品（手套、服装、床垫等）", "其他成品（毯子）"}
        )
        shown = {
            row["category_label"]
            for row in data["summary_rows"]
            if row["level"] == "category"
        }
        self.assertEqual(shown, {"其他成品"})
        # 科目名与摘要里的大类处理：科目带原值，摘要不带（辅助核算列已按要求去掉）
        self.assertEqual(
            {row["aux"] for block in data["blocks"] for row in block["entries"]
             if row["level"] == "line"}, {""}
        )
        # 借方只到渠道，不带销售大类（大类在贷方与附件里）
        self.assertTrue(
            all("其他成品" not in item for item in debit_accounts), debit_accounts
        )
        # 默认（fixed）口径仍按老规则归一，未受影响
        fixed = build_sales_voucher_summary(self.period)
        fixed_categories = {
            row["category"] for row in fixed["rows"] if row["level"] == "category"
        }
        self.assertIn("其他成品", fixed_categories)
        self.assertNotIn("其他成品（毯子）", fixed_categories)

    def test_voucher_period_filter_selects_month(self):
        """凭证期间筛选：选哪个期间就只出该期间的凭证。"""
        self._line(1, out_month=self.period.month, invoice_month=self.period.month,
                   carry_type="当月发货当月开票结转成本", cost="100.00", category="纱线")
        self._line(2, out_month=self.june, invoice_month=self.june,
                   carry_type="当月发货当月开票结转成本", cost="500.00", category="面料")
        current = build_sales_voucher_entries(self.period)
        self.assertEqual(current["month"], "2026-07")
        self.assertEqual(current["totals"]["debit"], Decimal("100.00"))
        june = build_sales_voucher_entries(self.period, month=self.june)
        self.assertEqual(june["month"], "2026-06")
        self.assertEqual(june["totals"]["debit"], Decimal("500.00"))
        # 页面上的凭证期间下拉含两个期间，且当前期间被选中
        self.client.force_login(self.user)
        html = self.client.get(
            reverse("sales_cost_voucher", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertIn("凭证期间", html)
        self.assertIn('name="voucher_month"', html)
        self.assertIn('value="2026-06"', html)
        self.assertIn('value="2026-07" selected', html)

    def test_voucher_export_listing(self):
        """销售结转成本凭证可单独导出 Excel。"""
        self._line(1, out_month=self.period.month, invoice_month=self.period.month,
                   carry_type="当月发货当月开票结转成本", cost="100.00",
                   category="纱线")
        self.client.force_login(self.user)
        response = self.client.get(
            reverse("export_listing", args=[self.period.id, "sales_voucher"])
        )
        self.assertEqual(response.status_code, 200)
        self.assertIn("spreadsheetml", response["Content-Type"])

    # ── 发出商品及销售成本统计表 ────────────────────────────────────────

    def _stat_line_set(self):
        # 5 月：发货 1000（未开票）、发货 2000（当月开票）
        self._line(11, out_month=self.may, carry_type="当月发货未开票挂发出商品",
                   status="uninvoiced", cost="1000.00")
        self._line(12, out_month=self.may, invoice_month=self.may,
                   carry_type="当月发货当月开票结转成本", cost="2000.00")
        # 6 月：上月的 1000 里开票结转 600；当月发货未开票 400；样品 50
        self._line(21, out_month=self.may, invoice_month=self.june,
                   carry_type="前期发货本期开票结转成本", cost="600.00")
        self._line(22, out_month=self.june, carry_type="当月发货未开票挂发出商品",
                   status="uninvoiced", cost="400.00")
        self._line(23, out_month=self.june, carry_type="免费样品赠送计入销售费用",
                   status="sample_free", cost="50.00")

    def test_statistics_rollforward_and_columns(self):
        self._stat_line_set()
        data = build_sales_cost_statistics(self.period, month_count=6)
        self.assertEqual(data["month_headers"],
                         ["2026年%d月" % m for m in range(7, 13)])
        cells = {month["value"]: month["cells"] for month in data["months"]}
        # 7 月：期初 = 5 月末余额（1000 未开票 − 600 开票结转）+ 6 月发货未开票 400 = 800
        self.assertEqual(cells["2026-07"]["上月期末发出商品未开票成本"], Decimal("800.00"))
        self.assertEqual(cells["2026-07"]["当月发货未开票挂发出商品"], Decimal("0.00"))
        self.assertEqual(cells["2026-07"]["当月期末发出商品未开票成本"], Decimal("800.00"))
        # 8-12 月：没有新数据，余额结转下去
        self.assertEqual(cells["2026-12"]["上月期末发出商品未开票成本"], Decimal("800.00"))
        self.assertEqual(cells["2026-12"]["当月期末发出商品未开票成本"], Decimal("800.00"))
        # 行结构：3 个一级分类 + 子项
        labels = [row["label"] for row in data["rows"] if row["level"] == "group"]
        self.assertEqual(labels, ["一、本月销售成本", "二、发出商品逐月结转", "三、其他"])
        items = [row["label"] for row in data["rows"] if row["level"] == "item"]
        self.assertEqual(items, [
            "当月发货当月开票结转成本", "前期发货本期开票结转成本", "本月主营业务成本合计",
            "上月期末发出商品未开票成本", "当月发货未开票挂发出商品",
            "前期发货本期开票结转成本", "当月期末发出商品未开票成本",
            "免费样品赠送计入销售费用",
        ])
        # 每个子项都有 6 个月份的单元格
        for row in data["rows"]:
            if row["level"] == "item":
                self.assertEqual(len(row["cells"]), 6, row["label"])

    def test_statistics_main_cost_total_and_sample(self):
        self._stat_line_set()
        data = build_sales_cost_statistics(self.period)
        cells = {month["value"]: month["cells"] for month in data["months"]}
        # 7 月：本测试数据里没有 7 月开票/7 月发货的结转，所以主营业务成本为 0
        self.assertEqual(cells["2026-07"]["当月发货当月开票结转成本"], Decimal("0.00"))
        self.assertEqual(cells["2026-07"]["前期发货本期开票结转成本"], Decimal("0.00"))
        self.assertEqual(cells["2026-07"]["本月主营业务成本合计"], Decimal("0.00"))
        # 6 月口径（以 6 月为起点才看得到）：前期发货本期开票 600 + 样品 50
        june_rows = build_sales_cost_statistics(
            AccountingPeriod(month=self.june), month_count=1
        )
        june_cell = june_rows["months"][0]["cells"]
        self.assertEqual(june_cell["前期发货本期开票结转成本"], Decimal("600.00"))
        self.assertEqual(june_cell["本月主营业务成本合计"], Decimal("600.00"))
        self.assertEqual(june_cell["免费样品赠送计入销售费用"], Decimal("50.00"))
        # 样品 50 记在 6 月（出库月），不在 7 月列里 → 7 月样品为 0
        self.assertEqual(cells["2026-07"]["免费样品赠送计入销售费用"], Decimal("0.00"))
        # 以 6 月为起点的月份列
        june_data = build_sales_cost_statistics(
            AccountingPeriod(month=self.june), month_count=2
        )
        self.assertEqual(june_data["month_headers"], ["2026年6月", "2026年7月"])

    def test_statistics_page_renders(self):
        self._stat_line_set()
        self.client.force_login(self.user)
        html = self.client.get(
            reverse("sales_uninvoiced", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertIn("<h2>发出商品及销售成本统计表</h2>", html)
        self.assertIn("序号", html)
        self.assertIn("项目分类", html)
        for label in ("一、本月销售成本", "二、发出商品逐月结转", "三、其他",
                      "上月期末发出商品未开票成本", "当月期末发出商品未开票成本"):
            self.assertIn(label, html)
        for label in ("2026年%d月" % m for m in range(7, 13)):
            self.assertIn(label, html)
        # 期末余额结转下去 = 800.00 应出现在页面上
        self.assertIn("800.00", html)


class SalesUninvoicedDetailLayoutTests(SettlementFixtureMixin, TestCase):
    """月末未开票页：明细表「出库月份/日期」拆成两列；明细上方那块筛选字段已删除。"""

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="uninv-layout", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.batch = self.create_batch(
            self.period, ImportBatch.Kind.SALES_LEDGER, "uninv-layout"
        )
        self.client.force_login(self.user)
        self.url = reverse("sales_uninvoiced", args=[self.period.id])

    def test_detail_table_uses_sales_ledger_columns(self):
        """明细表＝销售台账口径的列（出库期间 / 销售订单号 / 客户 … 销售成本）。"""
        html = self.client.get(self.url).content.decode("utf-8")
        idx = html.index(
            "<h2>截至月末仍未开票的销售出库明细（销售台账口径）</h2>"
        )
        section = html[html.rindex("<section", 0, idx):]
        section = section[: section.index("</section>")]
        heads = [re.sub(r"<[^>]+>", "", x).strip()
                 for x in re.findall(r"<th[^>]*>(.*?)</th>", section, re.S)]
        self.assertEqual(heads[0], "出库期间")
        self.assertEqual(heads[1], "销售订单号")
        self.assertIn("客户", heads)
        self.assertIn("销售成本", heads)
        self.assertIn("未开票天数", heads)
        self.assertNotIn("出库月份/日期", heads)
        # 每行列数与表头一致（含 colspan 累加）
        def width(tr):
            n = 0
            for td in re.findall(r"<td\b[^>]*>", tr):
                span = re.search(r'colspan="(\d+)"', td)
                n += int(span.group(1)) if span else 1
            return n
        body = section[section.index("<tbody>"):section.index("</tbody>")]
        for tr in re.findall(r"<tr\b[^>]*>(.*?)</tr>", body, re.S):
            self.assertEqual(width(tr), len(heads), tr[:80])

    def test_detail_filter_offers_period_and_customer_only(self):
        """明细表筛选（用户 2026-09-27）：只留「期间 / 客户」，并带导出与打印本表。"""
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertIn('name="issue_month"', html)
        self.assertIn('name="customer"', html)
        self.assertIn("导出Excel", html)
        self.assertIn('class="print-here"', html)
        # 旧筛选字段（此前已删除，不再回来）
        for field in ("document_no", "sales_order_no", "cross_period",
                      'name="status"', 'name="material"'):
            self.assertNotIn(field, html, field)

    # ── 新增：筛选 / 勾稽表 / 千分位 / 导出全部工作表（用户 2026-09-27） ──

    def _ledger_line(self, row_no, out_month, customer, carry_type, quantity, unit_cost):
        return SalesLedgerLine.objects.create(
            period=self.period,
            batch=self.batch,
            row_no=row_no,
            out_month=out_month,
            invoice_status=SalesLedgerLine.InvoiceStatus.UNINVOICED,
            carry_type=carry_type,
            customer_code="C%d" % row_no,
            customer_name=customer,
            material_code="FG-%d" % row_no,
            material_name="产品%d" % row_no,
            uom="kg",
            quantity=Decimal(quantity),
            cost_price=Decimal(unit_cost),
            cost_amount=(Decimal(quantity) * Decimal(unit_cost)).quantize(Decimal("0.01")),
        )

    def _check_rows(self, html):
        """解析「未开票 · 免费样品 · 发出商品跨月滚动勾稽」表的行（列表）。"""
        idx = html.index("未开票 · 免费样品 · 发出商品跨月滚动勾稽")
        section = html[html.rindex("<section", 0, idx):]
        section = section[: section.index("</section>")]
        body = section[section.index("<tbody>"): section.index("</tbody>")]
        rows = []
        for tr in re.findall(r"<tr\b[^>]*>(.*?)</tr>", body, re.S):
            cells = [
                re.sub(r"<[^>]+>", "", td).replace("&nbsp;", " ").strip()
                for td in re.findall(r"<td\b[^>]*>(.*?)</td>", tr, re.S)
            ]
            rows.append(cells)
        return rows

    def test_check_table_covers_uninvoiced_sample_and_rollforward(self):
        """勾稽表：未开票笔数 → 免费样品 → 发出商品跨月滚动，数字逐环相扣。

        方案A（用户 2026-09-27）：以前月份出库的那批＝「期初发出商品」，不再计入本期新增；
        这样本页期末发出商品与「发出商品及销售成本统计表」完全一致（差额 0）。
        """
        goods_batch = self.create_batch(
            self.period, ImportBatch.Kind.OPENING_GOODS_ISSUED, "uninv-check-opening"
        )
        # 期初发出商品（上传）：与下面 6 月那条未开票行是同一批货
        OpeningGoodsIssued.objects.create(
            period=self.period, batch=goods_batch,
            original_period=date(2026, 6, 1), document_date=date(2026, 6, 30),
            document_no="OG-CHECK", line_no="1", sales_order_no="SO-1",
            customer_name="客户甲", material_code="FG-1", material_name="产品1",
            uom="kg", original_quantity=Decimal("1000"),
            prior_invoiced_quantity=Decimal("0"), opening_quantity=Decimal("1000"),
            unit_cost=Decimal("2"), opening_amount=Decimal("2000.00"),
            remaining_quantity=Decimal("1000"),
        )
        self._ledger_line(1, date(2026, 6, 1), "客户甲", "当月发货未开票挂发出商品", "1000", "2")
        self._ledger_line(2, date(2026, 7, 1), "客户乙", "当月发货未开票挂发出商品", "500", "2")
        # 前期发货本期开票：开票月＝本期，才会滚进本期「减」项
        prior = self._ledger_line(
            3, date(2026, 6, 1), "客户丙", "前期发货本期开票结转成本", "300", "2"
        )
        prior.invoice_status = SalesLedgerLine.InvoiceStatus.INVOICED
        prior.invoice_month = date(2026, 7, 1)
        prior.save(update_fields=["invoice_status", "invoice_month"])
        self._ledger_line(4, date(2026, 7, 1), "客户丁", "免费样品赠送计入销售费用", "10", "2")
        html = self.client.get(self.url).content.decode("utf-8")
        rows = self._check_rows(html)
        # 模板给明细项加了 &nbsp; 缩进：取键时把空白（含 &nbsp;）全部去掉
        by_label = {
            re.sub(r"(&nbsp;|\s|\u3000|\u00a0)+", "", r[0]): r for r in rows if r
        }
        # 一、未开票笔数（＝明细表行数，2 笔：6月 1000 ＋ 7月 500）
        key = re.sub(r"(&nbsp;|\s|\u3000|\u00a0)+", "", "未开票笔数（＝下方明细表行数）")
        self.assertEqual(by_label[key][1], "2")
        self.assertEqual(by_label[key][3], "3,000.00")
        # 其中以前月份出库：6月那 1 笔 2,000.00
        prev_key = re.sub(r"(&nbsp;|\s|\u3000|\u00a0)+", "", "其中：以前月份出库、本月仍未开票")
        self.assertEqual(by_label[prev_key][1], "1")
        self.assertEqual(by_label[prev_key][3], "2,000.00")
        # 二、免费样品
        self.assertEqual(by_label["免费样品赠送计入销售费用"][1], "1")
        self.assertEqual(by_label["免费样品赠送计入销售费用"][3], "20.00")
        # 三、滚动：期初 2,000.00 ＋ 本期新增 1,000.00（只算 7 月出库那条）
        #            － 本期开票结转 600.00 ＝ 2,400.00
        self.assertEqual(by_label["期初发出商品"][3], "2,000.00")
        self.assertEqual(by_label["加：本期新增发出商品"][1], "1")
        self.assertEqual(by_label["加：本期新增发出商品"][3], "1,000.00")
        self.assertEqual(by_label["减：本期开票结转成本"][3], "600.00")
        end_row = [r for r in rows if r and r[0].startswith("＝")][0]
        self.assertEqual(end_row[3], "2,400.00")
        # 四、与统计表对账：两边都为 2,400.00 → 差额 0.00
        gap_row = [r for r in rows if r and r[0].startswith("差额")][0]
        self.assertEqual(gap_row[3], "0.00")

        # 本期新增不再包含以前月份出库那批（否则重复计量）
        from settlement.services.sales_uninvoiced import build_current_uninvoiced_rows
        current = build_current_uninvoiced_rows(self.period)
        self.assertEqual(len(current), 1)
        self.assertEqual(current[0].customer_name, "客户乙")

    def test_statistics_amounts_have_thousands_separator(self):
        """统计表金额加千分位（用户 2026-09-27）。"""
        self._ledger_line(1, date(2026, 7, 1), "客户甲", "当月发货未开票挂发出商品", "1000", "2")
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertIn('<td class="num">2,000.00</td>', html)
        self.assertNotIn('<td class="num">2000.00</td>', html)

    def test_detail_filter_and_export_share_the_same_rows(self):
        """筛选（期间 / 客户）与「导出Excel」同口径：页面行数 ＝ 导出行数。"""
        self._ledger_line(1, date(2026, 6, 1), "客户甲", "当月发货未开票挂发出商品", "1000", "2")
        self._ledger_line(2, date(2026, 7, 1), "客户乙", "当月发货未开票挂发出商品", "500", "2")

        page = self.client.get(self.url).content.decode("utf-8")
        self.assertIn("合计（2 行）", page)
        self.assertIn("共 2 笔（本期全部未开票行）", page)

        filtered = self.client.get(self.url, {"issue_month": "2026-06"}).content.decode("utf-8")
        self.assertIn("合计（1 行）", filtered)
        self.assertIn("共 1 笔（已按下方筛选条件过滤）", filtered)

        by_customer = self.client.get(self.url, {"customer": "客户乙"}).content.decode("utf-8")
        self.assertIn("合计（1 行）", by_customer)

        resp = self.client.get(
            reverse("export_listing", kwargs={"period_id": self.period.id, "kind": "sales_uninvoiced"}),
            {"issue_month": "2026-06"},
        )
        wb = load_workbook(BytesIO(resp.content), read_only=True)
        try:
            ws = wb["月末未开票"]
            rows = list(ws.iter_rows(min_row=2, values_only=True))
        finally:
            wb.close()
        self.assertEqual(len(rows), 1)

    def test_export_includes_check_and_statistics_sheets(self):
        """「导出Excel」＝整页内容：明细 ＋ 发出商品滚动勾稽 ＋ 统计表 三张工作表。"""
        self._ledger_line(1, date(2026, 7, 1), "客户甲", "当月发货未开票挂发出商品", "1000", "2")
        resp = self.client.get(
            reverse("export_listing", kwargs={"period_id": self.period.id, "kind": "sales_uninvoiced"})
        )
        wb = load_workbook(BytesIO(resp.content), read_only=True)
        try:
            names = list(wb.sheetnames)
            check = wb["发出商品滚动勾稽"]
            heads = [c.value for c in next(check.iter_rows(min_row=1, max_row=1))]
        finally:
            wb.close()
        self.assertIn("月末未开票", names)
        self.assertIn("发出商品滚动勾稽", names)
        self.assertIn("发出商品及销售成本统计表", names)
        self.assertEqual(heads, ["项目", "笔数", "数量", "金额", "说明"])

    def test_other_sections_still_present(self):
        html = self.client.get(self.url).content.decode("utf-8")
        for label in ("发出商品跨月滚动勾稽",
                      "发出商品及销售成本统计表", "截至月末仍未开票的销售出库明细"):
            self.assertIn(label, html)

    def test_voucher_section_moved_to_own_page(self):
        """销售结转成本凭证在独立页；本页既不展示凭证，也不再放跳转按钮
        （用户 2026-09-27 要求删掉页头那排按钮，凭证从导航进）。"""
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertNotIn("<h2>凭证分录（不含税）</h2>", html)
        self.assertNotIn("voucher-filter-form", html)
        self.assertNotIn(reverse("sales_cost_voucher", args=[self.period.id]), html)

        voucher_url = reverse("sales_cost_voucher", args=[self.period.id])
        voucher_html = self.client.get(voucher_url).content.decode("utf-8")
        self.assertIn("<h1>销售结转成本凭证</h1>", voucher_html)
        self.assertIn("<h2>凭证分录（不含税）</h2>", voucher_html)
        self.assertIn("voucher-filter-form", voucher_html)


class SalesVoucherDataSourceNoteTests(SettlementFixtureMixin, TestCase):
    """销售结转成本凭证（独立页面）：说明为「数据来源」，且不再显示旧的统计/筛选说明。"""

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="voucher-note", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.client.force_login(self.user)

    def test_note_shows_data_sources(self):
        html = self.client.get(
            reverse("sales_cost_voucher", args=[self.period.id])
        ).content.decode("utf-8")
        for text in (
            "数据来源",
            "销售台账",
            "结转类型",
            "销售成本",
            "销售数量 × 该产品月末加权平均价",
            "库存计价",
            "历史单位成本",
            "凭证期间（记账月份）",
            "有数量、金额为 0",
            "库存商品",
            "发出商品",
        ):
            self.assertIn(text, html, text)

    def test_old_note_is_removed(self):
        html = self.client.get(
            reverse("sales_cost_voucher", args=[self.period.id])
        ).content.decode("utf-8")
        for text in (
            "统计三类结转，按「结转类型 × 销售大类」汇总",
            "三个结转类型块的筛选各自独立",
            "来源台账",
        ):
            self.assertNotIn(text, html, text)

    def test_keeps_filters_and_summary_note(self):
        html = self.client.get(
            reverse("sales_cost_voucher", args=[self.period.id])
        ).content.decode("utf-8")
        # 表内筛选与合计说明仍在
        for text in ("记账月份", "（空白）", "小计", "合计", "默认筛选"):
            self.assertIn(text, html, text)


class PurchaseAccrualVoucherTests(SettlementFixtureMixin, TestCase):
    """采购暂估 / 暂估回冲凭证：三笔分录、价差、期初台账与页面渲染。"""

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="pv-tester", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.batch = self.create_batch(self.period, ImportBatch.Kind.INVENTORY, "pv")
        self.contract_batch = self.create_batch(
            self.period, ImportBatch.Kind.PURCHASE_CONTRACT, "pv-contract"
        )
        self.invoice_batch = self.create_batch(
            self.period, ImportBatch.Kind.PURCHASE_INVOICE, "pv-invoice"
        )
        self.material = Material.objects.create(
            code="RM-PV", name="暂估原料", uom="kg"
        )
        self.contract = PurchaseContractLine.objects.create(
            period=self.period,
            batch=self.contract_batch,
            contract_no="PC-PV",
            contract_date=date(2026, 7, 1),
            supplier_code="SUP-PV",
            supplier_name="暂估供应商",
            material_code=self.material.code,
            material_name=self.material.name,
            quantity=Decimal("100"),
            remaining_quantity=Decimal("100"),
            unit_price_excl_tax=Decimal("10"),
        )
        self.receipt = self.create_txn(
            self.period,
            self.batch,
            txn_date=date(2026, 7, 5),
            document_no="GR-PV",
            business_type="采购入库",
            direction=InventoryTransaction.Direction.IN,
            material=self.material,
            quantity=Decimal("100"),
            supplier_code="SUP-PV",
            supplier_name="暂估供应商",
            contract_no="PC-PV",
        )
        self.receipt.unit_price = Decimal("10")
        self.receipt.amount = Decimal("1000.00")
        self.receipt.price_source = "合同暂估"
        self.receipt.category_name = "原料"
        self.receipt.save()
        MatchAllocation.objects.create(
            period=self.period,
            kind=MatchAllocation.Kind.PURCHASE_CONTRACT,
            inventory_transaction=self.receipt,
            contract_line=self.contract,
            matched_quantity=Decimal("100"),
            unit_price=Decimal("10"),
            amount=Decimal("1000.00"),
        )

    def _match_invoice(self, *, invoice_no, quantity, unit_price, target=None):
        line = PurchaseInvoiceLine.objects.create(
            period=self.period,
            batch=self.invoice_batch,
            invoice_no=invoice_no,
            invoice_date=date(2026, 7, 20),
            invoice_period=date(2026, 7, 1),
            supplier_code="SUP-PV",
            supplier_name="暂估供应商",
            material_code=self.material.code,
            material_name=self.material.name,
            quantity=quantity,
            remaining_quantity=Decimal("0"),
            unit_price_excl_tax=unit_price,
            amount_excl_tax=(quantity * unit_price).quantize(Decimal("0.01")),
        )
        allocation = MatchAllocation.objects.create(
            period=self.period,
            kind=MatchAllocation.Kind.PURCHASE_INVOICE,
            purchase_invoice_line=line,
            matched_quantity=quantity,
            unit_price=unit_price,
            amount=(quantity * unit_price).quantize(Decimal("0.01")),
        )
        if target is not None:
            allocation.opening_purchase_accrual = target
        else:
            allocation.inventory_transaction = self.receipt
        allocation.save()
        return line

    def _opening_accrual(self, *, quantity, unit_price, remaining):
        batch = self.create_batch(
            self.period, ImportBatch.Kind.OPENING_PURCHASE_ACCRUAL, "pv-opening"
        )
        return OpeningPurchaseAccrual.objects.create(
            period=self.period,
            batch=batch,
            original_period=date(2026, 4, 1),
            document_date=date(2026, 4, 10),
            document_no="GR-OLD",
            supplier_code="SUP-PV",
            supplier_name="暂估供应商",
            material_code=self.material.code,
            material_name=self.material.name,
            original_quantity=quantity,
            opening_quantity=quantity,
            provisional_unit_price=unit_price,
            opening_amount=(quantity * unit_price).quantize(Decimal("0.01")),
            remaining_quantity=remaining,
        )

    def _extra_receipt(self, *, code, category, supplier, quantity, unit_price):
        material = Material.objects.create(code=code, name=code, uom="kg")
        txn = self.create_txn(
            self.period,
            self.batch,
            txn_date=date(2026, 7, 6),
            document_no="GR-%s" % code,
            business_type="采购入库",
            direction=InventoryTransaction.Direction.IN,
            material=material,
            quantity=quantity,
            supplier_code="SUP-%s" % code,
            supplier_name=supplier,
        )
        txn.unit_price = unit_price
        txn.amount = (quantity * unit_price).quantize(Decimal("0.01"))
        txn.price_source = "手工暂估"
        txn.category_name = category
        txn.save()
        return txn

    def test_same_month_invoice_splits_into_direct_and_accrual(self):
        """同月到货、只到票一部分：已到票走①、未到票走②，差额进材料成本差异。"""
        self._match_invoice(
            invoice_no="PI-PV-1", quantity=Decimal("60"), unit_price=Decimal("12")
        )
        data = build_purchase_voucher(self.period)
        summary = data["summary"]
        self.assertEqual(summary["receipt_amount"], Decimal("1000.00"))
        self.assertEqual(summary["same_month_amount"], Decimal("600.00"))
        self.assertEqual(summary["accrual_amount"], Decimal("400.00"))
        self.assertEqual(
            summary["same_month_amount"] + summary["accrual_amount"],
            summary["receipt_amount"],
        )
        self.assertEqual(summary["same_month_difference"], Decimal("120.00"))
        self.assertTrue(data["totals"]["balanced"])
        self.assertEqual(data["totals"]["difference"], Decimal("0.00"))

        first = {block["key"]: block for block in data["blocks"]}["same_month"]
        self.assertEqual(first["debit_total"], Decimal("720.00"))
        self.assertEqual(first["credit_total"], Decimal("720.00"))
        accounts = [row["account"] for row in first["rows"] if row["level"] == "line"]
        self.assertIn("材料成本差异", accounts)

    def test_reversal_block_clears_opening_accrual(self):
        """③ 前期暂估本月到票：先冲回暂估（借暂估应付/贷原材料），再按票入账（借原材料/贷应付账款）。"""
        opening = self._opening_accrual(
            quantity=Decimal("50"), unit_price=Decimal("8"), remaining=Decimal("0")
        )
        self._match_invoice(
            invoice_no="PI-PV-2", quantity=Decimal("50"), unit_price=Decimal("9"),
            target=opening,
        )
        data = build_purchase_voucher(self.period)
        summary = data["summary"]
        self.assertEqual(summary["reversal_amount"], Decimal("400.00"))
        self.assertEqual(summary["reversal_invoice_amount"], Decimal("450.00"))
        self.assertEqual(summary["reversal_difference"], Decimal("50.00"))
        self.assertEqual(summary["opening_count"], 1)
        self.assertTrue(data["totals"]["balanced"])

        block = {item["key"]: item for item in data["blocks"]}["reversal"]
        lines = [row for row in block["rows"] if row["level"] == "line"]
        # 冲回暂估：借 应付账款—暂估应付（汇总一行，不带供应商）400
        accrued = [row for row in lines if row["account"] == "应付账款—暂估应付"]
        self.assertEqual(len(accrued), 1)
        self.assertEqual(accrued[0]["debit"], Decimal("400.00"))
        self.assertEqual(accrued[0]["credit"], Decimal("0.00"))
        # 原材料两条腿：冲回时贷 400、按票入账时借 400（存货金额不因发票改写）
        raw = [row for row in lines if row["account"] == "原材料"]
        self.assertEqual(len(raw), 2, "原材料应借贷各一行（冲回暂估 + 按票入账）")
        self.assertEqual(
            sorted(row["debit"] for row in raw), [Decimal("0.00"), Decimal("400.00")]
        )
        self.assertEqual(
            sorted(row["credit"] for row in raw), [Decimal("0.00"), Decimal("400.00")]
        )
        # 按票入账：贷 应付账款 = 发票额 450
        supplier = [row for row in lines if row["account"] == "应付账款"]
        self.assertEqual(len(supplier), 1)
        self.assertEqual(supplier[0]["credit"], Decimal("450.00"))
        # 摘要措辞：冲回暂估 / 按票入账 两条腿都在「前期暂估本月到票」下
        summaries = {row["summary"] for row in lines}
        self.assertIn("前期暂估本月到票-冲回暂估", summaries)
        self.assertIn("前期暂估本月到票-按票入账", summaries)
        self.assertEqual(block["title"], "前期暂估本月到票")
        # 价差单独一行，进材料成本差异
        diff_rows = [row for row in block["rows"] if row["account"] == "材料成本差异"]
        self.assertEqual(len(diff_rows), 1)
        self.assertEqual(diff_rows[0]["debit"], Decimal("50.00"))
        # 借贷合计：借 400 + 400 + 50 = 贷 400 + 450
        self.assertEqual(block["debit_total"], Decimal("850.00"))
        self.assertEqual(block["credit_total"], Decimal("850.00"))

    def test_opening_balance_reported_but_not_posted(self):
        """期初暂估台账未核销：只提示与展示，不进凭证金额。"""
        self._opening_accrual(
            quantity=Decimal("50"), unit_price=Decimal("8"), remaining=Decimal("50")
        )
        data = build_purchase_voucher(self.period)
        summary = data["summary"]
        self.assertEqual(summary["opening_balance"], Decimal("400.00"))
        self.assertEqual(summary["current_balance"], Decimal("1000.00"))
        self.assertEqual(summary["ending_balance"], Decimal("1400.00"))
        # 凭证只包含本月到货 1000（未到票暂估），不含期初台账的 400
        self.assertEqual(data["totals"]["debit"], Decimal("1000.00"))
        self.assertTrue(
            any("期初暂估台账" in check["text"] for check in data["checks"])
        )

    def test_missing_price_is_blocking_check(self):
        """没有暂估单价的入库行：给出阻断提示，且不产生金额。

        用没有任何合同的新物料，避免「同物料合同价兜底」把价格补上。
        """
        material = Material.objects.create(code="RM-NOPRICE", name="无价原料", uom="kg")
        txn = self.create_txn(
            self.period,
            self.batch,
            txn_date=date(2026, 7, 7),
            document_no="GR-NOPRICE",
            business_type="采购入库",
            direction=InventoryTransaction.Direction.IN,
            material=material,
            quantity=Decimal("5"),
            supplier_code="SUP-PV",
            supplier_name="暂估供应商",
        )
        data = build_purchase_voucher(self.period)
        self.assertIn("blocking", [check["level"] for check in data["checks"]])
        self.assertTrue(
            any("没有暂估单价" in check["text"] for check in data["checks"])
        )
        statuses = [row["status"] for row in data["receipt_rows"]]
        self.assertIn("缺少暂估价格", statuses)
        self.assertEqual(data["summary"]["unpriced_rows"], 1)
        self.assertIsNone(txn.unit_price)

    def test_group_modes(self):
        """汇总模式合并成科目行；展开模式按物料大类 / 供应商分行。"""
        self._extra_receipt(
            code="RM-PV2", category="包装物", supplier="另一个供应商",
            quantity=Decimal("20"), unit_price=Decimal("5"),
        )
        single = build_purchase_voucher(self.period, group="single")
        category = build_purchase_voucher(self.period, group="category")
        # 科目都已汇总（原材料不分大类、应付账款不分供应商），两种模式结构一致
        self.assertEqual(
            single["totals"]["line_count"], category["totals"]["line_count"]
        )
        single_accounts = {
            row["account"]
            for block in single["blocks"]
            for row in block["rows"]
            if row["level"] == "line"
        }
        self.assertIn("原材料", single_accounts)
        self.assertIn("应付账款—暂估应付", single_accounts)
        category_accounts = {
            row["account"]
            for block in category["blocks"]
            for row in block["rows"]
            if row["level"] == "line"
        }
        self.assertIn("原材料", category_accounts)
        # 原材料不再按物料大类拆（原料 / 包装物 都不出现）
        self.assertNotIn("原材料—原料", category_accounts)
        self.assertNotIn("原材料—包装物", category_accounts)
        self.assertEqual(single["totals"]["debit"], category["totals"]["debit"])

    def test_purchase_voucher_period_filter(self):
        """凭证期间筛选：voucher_month 可切到别的核算期间，页面按该期间出凭证。"""
        june = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        self.client.force_login(self.user)
        url = reverse("purchase_accrual_voucher", args=[self.period.id])
        html = self.client.get(url).content.decode("utf-8")
        self.assertIn('name="voucher_month"', html)
        self.assertIn('value="2026-06"', html)
        self.assertIn('value="2026-07" selected', html)
        # 切到 6 月（该期间没有采购入库）：页面照常渲染，凭证为空
        june_html = self.client.get(
            url, {"voucher_month": "2026-06"}
        ).content.decode("utf-8")
        self.assertIn('value="2026-06" selected', june_html)
        self.assertEqual(june.id, june.id)  # 期间对象可用
        self.assertEqual(
            self.client.get(url, {"voucher_month": "2026-13"}).status_code, 200
        )

    def test_purchase_material_is_consolidated(self):
        """原材料不分物料大类（原料/纱线/面料…），汇总一行。"""
        self._match_invoice(
            invoice_no="PI-PV-11", quantity=Decimal("100"), unit_price=Decimal("10")
        )
        data = build_purchase_voucher(self.period, group="category")
        material_accounts = {
            row["account"]
            for block in data["blocks"]
            for row in block["rows"]
            if row["level"] == "line" and (row["account"] or "").startswith("原材料")
        }
        self.assertEqual(material_accounts, {"原材料"})

    def test_purchase_credit_is_consolidated(self):
        """应付账款类科目（含暂估应付）不按供应商展开，汇总一行。"""
        self._match_invoice(
            invoice_no="PI-PV-9", quantity=Decimal("100"), unit_price=Decimal("10")
        )
        data = build_purchase_voucher(self.period, group="category")
        for block in data["blocks"]:
            for row in block["rows"]:
                if row["level"] != "line" or not row["account"]:
                    continue
                self.assertNotIn("—神马", row["account"])
                if row["account"].startswith("应付账款"):
                    # 不带供应商后缀
                    self.assertIn(
                        row["account"],
                        {"应付账款", "应付账款—暂估应付"},
                        row["account"],
                    )
        self.assertTrue(data["totals"]["balanced"])

    def test_page_and_export(self):
        self._match_invoice(
            invoice_no="PI-PV-3", quantity=Decimal("100"), unit_price=Decimal("10")
        )
        self.client.force_login(self.user)
        response = self.client.get(
            reverse("purchase_accrual_voucher", args=[self.period.id])
        )
        html = response.content.decode("utf-8")
        for text in (
            "<h1>采购暂估 / 回冲凭证</h1>",
            "本月到货到票",
            "本月货到票未到-暂估入账",
            "前期暂估本月到票",
            "附件1：本月到货明细",
            "附件2：暂估回冲明细",
            "附件3：应付账款—暂估应付余额表",
            "暂估账龄（按到货月份）",
            "凭证期间",
        ):
            self.assertIn(text, html, text)
        self.assertIn("1,000.00", html)
        # 与销售结转成本凭证同格式：无汇总卡、无借贷 pill、5 列分录
        self.assertNotIn("summary-strip", html)
        self.assertNotIn("借贷平衡", html)
        idx = html.index("<h2>凭证分录（不含税）</h2>")
        section = html[idx:html.index("</table>", idx)]
        heads = [
            re.sub(r"<[^>]+>", "", h).strip()
            for h in re.findall(r"<th[^>]*>(.*?)</th>", section, re.S)
        ]
        self.assertEqual(heads, ["摘要", "科目", "数量", "借方", "贷方"])

        export = self.client.get(
            reverse("export_listing", args=[self.period.id, "purchase_voucher"])
        )
        self.assertEqual(export.status_code, 200)
        workbook = load_workbook(BytesIO(export.content))
        self.assertEqual(workbook.sheetnames[0], "凭证分录")
        header_row = [
            cell for cell in next(workbook["凭证分录"].iter_rows(values_only=True))
        ]
        self.assertEqual(
            list(header_row),
            ["凭证期间", "摘要", "科目", "数量", "借方", "贷方"],
        )
        for name in ("附件1-本月到货明细", "附件2-暂估回冲明细",
                     "附件3-应付暂估余额", "附件4-暂估账龄"):
            self.assertIn(name, workbook.sheetnames)

    def test_filter_only_affects_attachments(self):
        """附件筛选不改变凭证金额。"""
        full = build_purchase_voucher(self.period)
        filtered = build_purchase_voucher(self.period, detail_supplier="不存在的供应商")
        self.assertEqual(full["totals"], filtered["totals"])
        self.assertEqual(filtered["receipt_rows"], [])
        self.assertTrue(filtered["filter"]["active"])


class MaterialIssueVoucherTests(SettlementFixtureMixin, TestCase):
    """生产领料（投料）凭证：除浆粕外按工单投料台账，浆粕按流水备注，其余归无工单领料。"""

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="mi-tester", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.batch = self.create_batch(self.period, ImportBatch.Kind.INVENTORY, "mi")
        self.raw = Material.objects.create(code="RM-MI", name="示例原料", uom="kg")
        self.yarn = Material.objects.create(code="11.01-MI", name="示例纱线", uom="kg")

    def _order(self, order_no, dept):
        return WorkOrder.objects.create(
            period=self.period, order_no=order_no, production_dept_code=dept
        )

    def _price(self, material, price):
        return InventoryValuation.objects.create(
            period=self.period,
            material_code=material.code,
            material_name=material.name,
            weighted_average_price=Decimal(price),
        )

    def _issue(self, *, work_order_no, material, quantity, returned="0",
               category="原料"):
        actual = Decimal(quantity) - Decimal(returned)
        return MaterialIssue.objects.create(
            period=self.period,
            batch=self.batch,
            work_order_no=work_order_no,
            material_code=material.code,
            material_name=material.name,
            material_category=category,
            issue_quantity=Decimal(quantity),
            returned_quantity=Decimal(returned),
            actual_issue_quantity=actual,
            issue_time=datetime(2026, 7, 10, 8, 0),
        )

    def _pick(self, *, document_no, material, quantity, amount, category,
              remark="", warehouse="临海主仓库"):
        txn = self.create_txn(
            self.period,
            self.batch,
            txn_date=date(2026, 7, 31),
            document_no=document_no,
            business_type="领料出库",
            direction=InventoryTransaction.Direction.OUT,
            material=material,
            quantity=Decimal(quantity),
        )
        txn.category_name = category
        txn.warehouse = warehouse
        txn.remark = remark
        txn.unit_price = (
            (Decimal(amount) / Decimal(quantity)) if Decimal(quantity) else None
        )
        txn.amount = Decimal(amount)
        txn.save()
        return txn

    def _fixture(self):
        """染色/纺纱走工单投料台账；浆粕走流水备注；其余进无工单领料。"""
        self._order("WO-DYE", "RanZheng")
        self._order("WO-SPIN", "tefang")
        self._price(self.raw, "10")
        self._price(self.yarn, "50")
        # 工单投料台账：染色领原料 100（1000）+ 纱线 10（500）；纺纱领原料 20 退 5 → 实际 15（150）
        self._issue(work_order_no="WO-DYE", material=self.raw, quantity="100")
        self._issue(work_order_no="WO-DYE", material=self.yarn, quantity="10",
                    category="纱线")
        self._issue(work_order_no="WO-SPIN", material=self.raw, quantity="20",
                    returned="5")
        # 领料出库：备注含「浆粕+7月」→ 浆粕车间；备注写浆粕但 6 月 → 无工单领料；别家备注 → 无工单领料
        self._pick(document_no="OUT-1", material=self.raw, quantity="7",
                   amount="70.00", category="原料", remark="浆粕车间7月领用",
                   warehouse="临海浆粕仓库")
        self._pick(document_no="OUT-2", material=self.raw, quantity="5",
                   amount="50.00", category="原料", remark="浆粕车间6月领用",
                   warehouse="临海浆粕仓库")
        self._pick(document_no="OUT-3", material=self.yarn, quantity="8",
                   amount="80.00", category="面料", remark="郝工制衣用布")

    def test_ledger_drives_non_pulp_workshops(self):
        self._fixture()
        data = build_material_issue_voucher(self.period)
        self.assertEqual(data["headers"], ["摘要", "科目", "数量", "借方", "贷方"])

        # 一张凭证：只有一条分录流，没有分块、没有小计
        self.assertNotIn("blocks", data)
        self.assertEqual([row["level"] for row in data["rows"]].count("total"), 1)
        self.assertNotIn("subtotal", [row["level"] for row in data["rows"]])
        line_rows = [row for row in data["rows"] if row["level"] == "line"]

        # 借方合并一行
        debit_rows = [row for row in line_rows if row["debit"]]
        self.assertEqual(len(debit_rows), 1)
        self.assertEqual(debit_rows[0]["account"], "生产成本—直接材料")
        self.assertEqual(debit_rows[0]["summary"], "7月各车间领用材料")
        # 台账 1000 + 500 + 150 = 1650；浆粕 70；无工单 50 + 80 = 130
        self.assertEqual(debit_rows[0]["debit"], Decimal("1850.00"))

        # 贷方按车间顺序：染色(原料/纱线) → 纺纱(原料) → 浆粕(原料) → 无工单(原料/面料)
        self.assertEqual(
            [row["summary"] for row in line_rows if row["credit"]],
            [
                "7月染色车间领用原料",
                "7月染色车间领用纱线",
                "7月纺纱车间领用原料",
                "7月浆粕车间领用原料",
                "7月无工单领料（原料）",
                "7月无工单领料（面料）",
            ],
        )
        credit_map = {
            row["summary"]: (row["account"], row["credit"])
            for row in line_rows if row["credit"]
        }
        self.assertEqual(credit_map["7月染色车间领用原料"],
                         ("原材料—原料", Decimal("1000.00")))
        self.assertEqual(credit_map["7月染色车间领用纱线"],
                         ("库存商品—纱线", Decimal("500.00")))
        self.assertEqual(credit_map["7月纺纱车间领用原料"],
                         ("原材料—原料", Decimal("150.00")))
        self.assertEqual(credit_map["7月浆粕车间领用原料"],
                         ("原材料—原料", Decimal("70.00")))
        self.assertEqual(credit_map["7月无工单领料（原料）"],
                         ("原材料—原料", Decimal("50.00")))
        self.assertEqual(credit_map["7月无工单领料（面料）"],
                         ("库存商品—面料", Decimal("80.00")))

        self.assertTrue(data["totals"]["balanced"])
        self.assertEqual(data["totals"]["debit"], Decimal("1850.00"))
        self.assertEqual(data["totals"]["credit"], Decimal("1850.00"))

        # 附件3 按科目跨车间合并
        self.assertEqual(
            {row["account"]: row["amount"] for row in data["credit_rows"]},
            {
                "原材料—原料": Decimal("1270.00"),
                "库存商品—纱线": Decimal("500.00"),
                "库存商品—面料": Decimal("80.00"),
            },
        )

        # 附件2：车间 × 物料大类
        detail = {
            (row["workshop_label"], row["category"]): row
            for row in data["category_rows"]
        }
        self.assertEqual(detail[("染色车间", "原料")]["amount"], Decimal("1000.00"))
        self.assertEqual(detail[("染色车间", "纱线")]["amount"], Decimal("500.00"))
        # 台账的实际投料数量已扣回撤：20 − 5 = 15
        self.assertEqual(detail[("纺纱车间", "原料")]["amount"], Decimal("150.00"))
        self.assertEqual(detail[("纺纱车间", "原料")]["quantity"], Decimal("15.000000"))
        self.assertEqual(detail[("浆粕车间", "原料")]["amount"], Decimal("70.00"))
        self.assertEqual(
            detail[("无工单领料", "原料")]["amount"] + detail[("无工单领料", "面料")]["amount"],
            Decimal("130.00"),
        )
        self.assertEqual(
            sum((row["amount"] for row in data["category_rows"]), Decimal("0")),
            Decimal("1850.00"),
        )

        # 汇总与检查
        self.assertEqual(data["summary"]["ledger_rows"], 3)
        self.assertEqual(data["summary"]["ledger_amount"], Decimal("1650.00"))
        self.assertEqual(data["summary"]["pulp_rows"], 1)
        self.assertEqual(data["summary"]["pulp_amount"], Decimal("70.00"))
        self.assertEqual(data["summary"]["mismatch_rows"], 1)
        self.assertEqual(data["summary"]["mismatch_amount"], Decimal("50.00"))
        self.assertEqual(data["summary"]["unallocated_rows"], 2)
        self.assertEqual(data["summary"]["unallocated_amount"], Decimal("130.00"))
        self.assertTrue(any("无工单领料" in c["text"] for c in data["checks"]))
        self.assertTrue(any("月份不是当期" in c["text"] for c in data["checks"]))
        self.assertTrue(
            any("浆粕车间领料" in c["text"] and c["level"] == "info"
                for c in data["checks"])
        )

        # 明细来源：3 条台账 + 3 条流水
        sources = [row["source"] for row in data["detail_rows"]]
        self.assertEqual(sources.count("工单投料台账"), 3)
        self.assertEqual(sources.count("库存领料出库"), 3)

    def test_zero_price_ledger_row_is_blocking(self):
        """台账物料没有月末加权价 → 金额按 0 计，必须阻断提示。"""
        self._order("WO-DYE", "RanZheng")
        self._issue(work_order_no="WO-DYE", material=self.raw, quantity="10")
        data = build_material_issue_voucher(self.period)
        self.assertEqual(data["summary"]["ledger_amount"], Decimal("0.00"))
        self.assertTrue(
            any(c["level"] == "blocking" and "工单投料台账" in c["text"]
                for c in data["checks"])
        )

    def test_flow_missing_from_ledger_is_reported(self):
        """投料流水在台账里找不到对应记录 → 不计入，但要在检查里说明。"""
        self._order("WO-DYE", "RanZheng")
        self._price(self.raw, "10")
        self._issue(work_order_no="WO-DYE", material=self.raw, quantity="100")
        txn = self.create_txn(
            self.period,
            self.batch,
            txn_date=date(2026, 7, 12),
            document_no="WO-DYE",
            business_type="投料",
            direction=InventoryTransaction.Direction.OUT,
            material=self.raw,
            quantity=Decimal("2"),
        )
        txn.category_name = "原料"
        txn.amount = Decimal("20.00")
        txn.save()
        data = build_material_issue_voucher(self.period)
        # 凭证只按台账 100×10=1000
        self.assertEqual(data["totals"]["debit"], Decimal("1000.00"))
        self.assertTrue(
            any("找不到对应记录" in c["text"] for c in data["checks"])
        )

    def test_page_center_and_export(self):
        self._fixture()
        self.client.force_login(self.user)
        url = reverse("material_issue_voucher", args=[self.period.id])
        html = self.client.get(url).content.decode("utf-8")
        self.assertIn("<h1>生产领料（投料）凭证</h1>", html)
        # 不分块：页面上不应该再出现 ①②③ 的车间分块标题，也没有分块小计行
        self.assertNotIn("① 染色车间", html)
        self.assertNotIn("车间 小计", html)
        self.assertIn("7月染色车间领用原料", html)
        self.assertIn("7月各车间领用材料", html)
        self.assertIn("工单投料台账", html)
        self.assertIn("无工单领料", html)
        self.assertIn("原材料—原料", html)
        self.assertIn("附件1：领料明细", html)
        self.assertIn("附件3：贷方科目汇总", html)
        self.assertIn("1,850.00", html)

        center = self.client.get(
            reverse("voucher_center", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertIn('id="voucher-material_issue"', center)
        self.assertIn(url, center)
        self.assertIn("7月各车间领用材料", center)
        self.assertIn("7月染色车间领用原料", center)

        listing = reverse(
            "export_listing", args=[self.period.id, "material_issue_voucher"]
        )
        self.assertEqual(self.client.get(listing).status_code, 200)
        resp = self.client.post(
            reverse("export_center", args=[self.period.id]),
            {"reports": ["material_issue_voucher"]},
        )
        self.assertEqual(resp.status_code, 200)

    def test_voucher_month_filter_switches_period(self):
        self._fixture()
        AccountingPeriod.objects.create(month=date(2026, 6, 1))
        self.client.force_login(self.user)
        url = reverse("material_issue_voucher", args=[self.period.id])
        html = self.client.get(
            url, {"voucher_month": "2026-06"}
        ).content.decode("utf-8")
        # 2026-06 没有台账与领料流水：分录只有合计行（页头/口径说明里的示例文字不算）
        self.assertNotIn("<td>7月各车间领用材料</td>", html)
        self.assertNotIn("<td>7月染色车间领用原料</td>", html)
        self.assertIn("本期没有工单投料台账", html)


class MaterialIssueSummaryTests(SettlementFixtureMixin, TestCase):
    """工单投料：汇总表按「投料期间 × 投料车间 × 物料分类」给出数量/金额/无单价行数。"""

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="om-tester", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.batch = self.create_batch(self.period, ImportBatch.Kind.INVENTORY, "om")
        self.raw = Material.objects.create(code="RM-OM", name="汇总原料", uom="kg")
        self.yarn = Material.objects.create(code="YM-OM", name="汇总纱线", uom="kg")
        WorkOrder.objects.create(
            period=self.period, order_no="WO-OM", production_dept_code="RanZheng"
        )
        InventoryValuation.objects.create(
            period=self.period,
            material_code=self.raw.code,
            material_name=self.raw.name,
            weighted_average_price=Decimal("10"),
        )

    def _issue(self, material, quantity, category, when):
        return MaterialIssue.objects.create(
            period=self.period,
            batch=self.batch,
            work_order_no="WO-OM",
            material_code=material.code,
            material_name=material.name,
            material_category=category,
            issue_quantity=Decimal(quantity),
            returned_quantity=Decimal("0"),
            actual_issue_quantity=Decimal(quantity),
            issue_time=when,
        )

    def test_summary_groups_by_period_workshop_category(self):
        from settlement.views import build_material_issue_rows, build_material_issue_summary

        self._issue(self.raw, "100", "原料", datetime(2026, 7, 5, 9, 0))
        self._issue(self.yarn, "5", "纱线", datetime(2026, 7, 6, 9, 0))
        # 投料时间缺失 → 期间退回核算期间
        self._issue(self.raw, "3", "原料", None)

        rows, totals = build_material_issue_rows(self.period)
        summary, summary_totals = build_material_issue_summary(rows, self.period)

        self.assertEqual([item["month"] for item in summary], ["2026-07"] * 2)
        self.assertEqual(
            [(item["workshop_label"], item["category"]) for item in summary],
            [("染色车间", "原料"), ("染色车间", "纱线")],
        )
        raw_row = summary[0]
        self.assertEqual(raw_row["quantity"], Decimal("103.000000"))
        self.assertEqual(raw_row["amount"], Decimal("1030.00"))
        self.assertEqual(raw_row["no_price"], 0)
        self.assertEqual(raw_row["rows"], 2)
        yarn_row = summary[1]
        self.assertEqual(yarn_row["quantity"], Decimal("5.000000"))
        self.assertEqual(yarn_row["amount"], Decimal("0.00"))
        self.assertEqual(yarn_row["no_price"], 1)

        self.assertEqual(summary_totals["rows"], 3)
        self.assertEqual(summary_totals["quantity"], Decimal("108.000000"))
        self.assertEqual(summary_totals["amount"], Decimal("1030.00"))
        self.assertEqual(summary_totals["no_price"], 1)

    def test_pulp_workshop_included_via_flow_remark(self):
        """浆粕车间没有工单：按库存流水的备注认领，并入同一张汇总表。"""
        from settlement.services.material_issue_voucher import pulp_issue_rows
        from settlement.views import build_material_issue_rows, build_material_issue_summary

        self._issue(self.raw, "100", "原料", datetime(2026, 7, 5, 9, 0))
        # 浆粕车间领料：备注含「浆粕 + 7月」→ 计入；备注写 6 月 → 不计入
        for document_no, remark, amount in (
            ("OUT-P1", "浆粕车间7月领用", "200.00"),
            ("OUT-P2", "浆粕车间6月领用", "300.00"),
        ):
            txn = self.create_txn(
                self.period,
                self.batch,
                txn_date=date(2026, 7, 31),
                document_no=document_no,
                business_type="领料出库",
                direction=InventoryTransaction.Direction.OUT,
                material=self.raw,
                quantity=Decimal("20"),
            )
            txn.category_name = "原料"
            txn.warehouse = "临海浆粕仓库"
            txn.remark = remark
            txn.amount = Decimal(amount)
            txn.save()

        pulp = pulp_issue_rows(self.period)
        self.assertEqual(len(pulp), 1)
        self.assertEqual(pulp[0]["category"], "原料")

        rows, _ = build_material_issue_rows(self.period)
        extra = [
            {
                "month": r["date"].strftime("%Y-%m"),
                "workshop": r["workshop"],
                "category": r["category"],
                "quantity": r["quantity"],
                "amount": r["amount"],
                "no_price": r["unit_price"] is None,
            }
            for r in pulp
        ]
        summary, totals = build_material_issue_summary(rows, self.period, extra_rows=extra)
        pulp_row = [item for item in summary if item["workshop"] == "浆粕车间"]
        self.assertEqual(len(pulp_row), 1)
        self.assertEqual(pulp_row[0]["category"], "原料")
        self.assertEqual(pulp_row[0]["quantity"], Decimal("20.000000"))
        self.assertEqual(pulp_row[0]["amount"], Decimal("200.00"))
        # 台账 1 行 + 浆粕 1 行
        self.assertEqual(totals["rows"], 2)
        self.assertEqual(totals["ledger_rows"], 1)
        self.assertEqual(totals["extra_rows"], 1)
        self.assertEqual(totals["amount"], Decimal("1200.00"))

        # 页面：浆粕车间出现在汇总表里，也作为一个页签进入明细表，明细里写来源与取数依据
        self.client.force_login(self.user)
        html = self.client.get(
            reverse("order_material_report", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertIn("浆粕车间", html)
        self.assertIn("浆粕车间按流水备注", html)
        self.assertIn("1,200.00", html)
        # 浆粕明细行：单据号写在「工单编号 / 单据号」列，来源标注
        self.assertIn("OUT-P1", html)
        self.assertIn("库存领料出库", html)
        # 明细表新增两列
        self.assertIn("<th>投料期间</th><th>工单编号 / 单据号</th><th>物料分类</th>", html)
        # 用户 2026-09-21：明细再补 6 列（批次号 / 投料仓位 / 投料时间 / 投料人员 /
        # 主产出物料编号 / 主产出物料名称）→ 表头 16 列
        seg = html.split("工单投料（按投料车间）</h2>", 1)[1].split("</table>", 1)[0]
        heads = [re.sub(r"<[^>]+>", "", h).strip()
                 for h in re.findall(r"<th[^>]*>(.*?)</th>", seg, re.S)]
        self.assertEqual(
            heads,
            ["投料期间", "工单编号 / 单据号", "物料分类", "物料编码", "物料名称", "规格",
             "单位", "实际投料数量", "批次号", "投料仓位", "投料时间", "投料人员",
             "主产出物料编号", "主产出物料名称", "单价", "金额"],
        )
        # 每行 16 格（含浆粕那张表）
        body = seg.split("</thead>", 1)[1]
        for row in re.findall(r"<tr>\s*(.*?)</tr>", body, re.S):
            self.assertEqual(len(re.findall(r"<td[^>]*>", row)), 16, row[:60])

    def test_page_renders_summary_table(self):
        from settlement.views import build_material_issue_rows, build_material_issue_summary

        self._issue(self.raw, "100", "原料", datetime(2026, 7, 5, 9, 0))
        self._issue(self.yarn, "5", "纱线", datetime(2026, 7, 6, 9, 0))
        rows, _ = build_material_issue_rows(self.period)
        summary, summary_totals = build_material_issue_summary(rows, self.period)
        self.assertEqual(len(summary), 2)
        self.assertEqual(summary_totals["no_price"], 1)

        self.client.force_login(self.user)
        resp = self.client.get(reverse("order_material_report", args=[self.period.id]))
        html = resp.content.decode("utf-8")
        self.assertEqual(resp.status_code, 200)
        for text in (
            "<h1>工单投料</h1>",
            "投料汇总（按投料期间 × 投料车间 × 物料分类）",
            "<section class=\"ws-tabs\" role=\"tablist\" data-no-print>",
            "<th>投料期间</th><th>工单编号 / 单据号</th><th>物料分类</th>",
            '无单价行数',
            "染色车间",
            "投料行数合计 2 行",
            "工单投料台账 2 行",
            "无单价 1 行",
        ):
            self.assertIn(text, html, text)
        # 旧的 4 张汇总卡（投料行数 / 实际投料数量 / 投料金额 / 无单价行数）不再出现在页面上
        self.assertNotIn("投料行数</span>", html)

        # 用户 2026-09-21：车间页签里不再显示「N 行」；每个车间表在「打印本表」旁加「导出Excel」
        tab_labels = re.findall(r'<label class="ws-tab"[^>]*>(.*?)</label>', html, re.S)
        self.assertTrue(tab_labels)
        for label in tab_labels:
            self.assertNotIn("行", re.sub(r"<[^>]+>", "", label))
        self.assertNotIn("行</small>", html)
        # 每个车间面板：标题 + 导出Excel（带该车间 workshop）+ print-here（打印本表落点，在导出右边）
        panels = html.split('<div class="ws-panel"')[1:]
        self.assertTrue(panels)
        for panel in panels:
            head = panel.split("</thead>", 1)[0]
            self.assertIn('class="print-here"', head)
            href = re.search(
                r'href="([^"]*export-list/order_material/[^"]*)"', head
            )
            self.assertIsNotNone(href)
            self.assertIn("workshop=", href.group(1))
            self.assertLess(
                head.index("导出Excel"), head.index('class="print-here"')
            )
        # 页头另有一个「导出Excel（全部车间）」（不带 workshop）
        page_head = html.split("</section>", 1)[0]
        self.assertIn("导出Excel（全部车间）", page_head)
        self.assertNotIn("workshop=", page_head)

    def test_workshop_export_only_contains_that_workshop(self):
        """每个车间的「导出Excel」只导该车间的行（?workshop= 生效）。"""
        from settlement.services.imports import normalize_workshop

        self.client.force_login(self.user)
        self._issue(self.raw, "100", "原料", datetime(2026, 7, 5, 9, 0))
        code = normalize_workshop("RanZheng")
        resp = self.client.get(
            reverse("export_listing", args=[self.period.id, "order_material"]),
            {"workshop": code},
        )
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(BytesIO(resp.content), read_only=True)
        try:
            ws = wb[wb.sheetnames[0]]
            heads = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
            rows = list(ws.iter_rows(min_row=2, values_only=True))
        finally:
            wb.close()
        self.assertIn("工单编号", heads)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][heads.index("工单编号")], "WO-OM")
        # 换成别的车间（本测试里没有那个车间的投料）→ 0 行
        resp2 = self.client.get(
            reverse("export_listing", args=[self.period.id, "order_material"]),
            {"workshop": normalize_workshop("BaoFu")},
        )
        wb2 = load_workbook(BytesIO(resp2.content), read_only=True)
        try:
            self.assertEqual(sum(1 for _ in wb2[wb2.sheetnames[0]].iter_rows(min_row=2)), 0)
        finally:
            wb2.close()


class MaterialIssueFilterTests(SettlementFixtureMixin, TestCase):
    """工单投料页面的筛选：分别放在「投料汇总」与「工单投料（按投料车间）」两个模块里，各筛各的；筛选栏单独一行靠左。"""

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="om-filter", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.batch = self.create_batch(self.period, ImportBatch.Kind.INVENTORY, "omf")
        self.raw = Material.objects.create(code="RM-F", name="筛选原料", uom="kg")
        self.yarn = Material.objects.create(code="YM-F", name="筛选纱线", uom="kg")
        WorkOrder.objects.create(
            period=self.period, order_no="WO-F-1", production_dept_code="RanZheng"
        )
        WorkOrder.objects.create(
            period=self.period, order_no="WO-F-2", production_dept_code="tefang"
        )
        InventoryValuation.objects.create(
            period=self.period, material_code=self.raw.code,
            material_name=self.raw.name, weighted_average_price=Decimal("10"),
        )
        InventoryValuation.objects.create(
            period=self.period, material_code=self.yarn.code,
            material_name=self.yarn.name, weighted_average_price=Decimal("20"),
        )
        # 染色：原料 100×10；纺纱：纱线 5×20
        MaterialIssue.objects.create(
            period=self.period, batch=self.batch, work_order_no="WO-F-1",
            material_code=self.raw.code, material_name=self.raw.name,
            material_category="原料", issue_quantity=Decimal("100"),
            returned_quantity=Decimal("0"), actual_issue_quantity=Decimal("100"),
            issue_time=datetime(2026, 7, 5, 8, 0),
        )
        MaterialIssue.objects.create(
            period=self.period, batch=self.batch, work_order_no="WO-F-2",
            material_code=self.yarn.code, material_name=self.yarn.name,
            material_category="纱线", issue_quantity=Decimal("5"),
            returned_quantity=Decimal("0"), actual_issue_quantity=Decimal("5"),
            issue_time=datetime(2026, 7, 6, 8, 0),
        )

    def _html(self, **params):
        self.client.force_login(self.user)
        resp = self.client.get(
            reverse("order_material_report", args=[self.period.id]), params
        )
        self.assertEqual(resp.status_code, 200)
        return resp.content.decode("utf-8")

    def _summary_block(self, html):
        start = html.index("投料汇总（按投料期间")
        end = html.index('<section class="ws-tabs"', start)
        return html[start:end]

    def _detail_block(self, html):
        start = html.index('<section class="ws-tabs"')
        return html[start:]

    def test_summary_module_has_no_export_or_print_button(self):
        """「投料汇总」这块自己不放导出/打印（靠 data-no-print 跳过自动按钮）。

        注意（用户 2026-09-21）：页面其它位置现在有导出按钮 ——
        页头「导出Excel（全部车间）」+ 每个车间页签里各自的「导出Excel」，
        所以这里只断言**汇总板块内部**没有导出/打印。
        """
        html = self._html()
        summary = self._summary_block(html)
        self.assertTrue(summary.lstrip().startswith('<section class="section-card" data-no-print>')
                        or '<section class="section-card" data-no-print>' in html)
        self.assertNotIn("导出Excel", summary)
        self.assertNotIn("打印本表", summary)
        # 明细模块的 6 张车间表仍由脚本自动加「打印本表」，页签容器自身跳过
        self.assertIn('<section class="ws-tabs" role="tablist" data-no-print>', html)

    def test_filters_live_in_each_module(self):
        """两个模块各自带筛选栏，没有单独的筛选模块。"""
        html = self._html()
        self.assertNotIn('<section class="section-card filter-card">', html)
        summary = self._summary_block(html)
        detail = self._detail_block(html)
        # 两个模块的筛选栏都在标题下面、单独一行靠左（不再挂在标题右侧）
        self.assertEqual(html.count('class="om-filter-row"'), 2)
        self.assertNotIn('style="margin:0;"', html)
        # 汇总模块：期间 / 车间 / 物料分类
        for name in ("s_month", "s_workshop", "s_category"):
            self.assertIn('name="%s"' % name, summary, name)
        self.assertNotIn('name="s_order_no"', summary)
        # 明细模块：期间 / 车间 / 物料分类 / 工单号 / 物料
        for name in ("d_month", "d_workshop", "d_category", "d_order_no", "d_material"):
            self.assertIn('name="%s"' % name, detail, name)
        # 两边都保留对方的条件（提交后另一模块状态不丢）
        self.assertIn('name="d_category"', summary)
        self.assertIn('name="s_category"', detail)

    def test_summary_filter_only_affects_summary(self):
        """汇总模块筛「纱线」：汇总只剩纺纱车间，明细表仍是两个车间。"""
        html = self._html(s_category="纱线")
        summary = self._summary_block(html)
        detail = self._detail_block(html)
        self.assertNotIn("<td>染色车间</td>", summary)
        self.assertIn("<td>纺纱车间</td>", summary)
        # 明细表不受影响：两个车间页签都在
        self.assertIn("纺纱车间", detail)
        self.assertIn("染色车间", detail)
        self.assertIn("本表已按上方条件过滤", summary)
        self.assertNotIn("明细表已按上方条件过滤", html)

    def test_detail_filter_only_affects_detail(self):
        """明细模块筛「纱线」：明细只剩纺纱车间，汇总仍是两行。"""
        html = self._html(d_category="纱线")
        summary = self._summary_block(html)
        detail = self._detail_block(html)
        self.assertIn("<td>染色车间</td>", summary)      # 汇总不受影响
        self.assertIn("<td>纺纱车间</td>", summary)
        self.assertNotIn("<td>染色车间</td>", detail)
        self.assertIn("纺纱车间", detail)
        self.assertIn("明细表已按上方条件过滤", detail)

    def test_detail_text_filters(self):
        html = self._html(d_order_no="WO-F-1")
        detail = self._detail_block(html)
        self.assertIn("WO-F-1", detail)
        self.assertNotIn("WO-F-2", detail)
        html = self._html(d_material="YM-F")
        detail = self._detail_block(html)
        self.assertIn("筛选纱线", detail)
        self.assertNotIn("筛选原料", detail)
        # 投料期间筛到别的月份：两边都空，但两个模块的筛选栏都还在（可清空恢复）
        html = self._html(s_month="2026-06", d_month="2026-06")
        self.assertIn("没有符合条件的投料记录", self._summary_block(html))
        detail = self._detail_block(html)
        self.assertIn("没有符合条件的投料记录", detail)
        for name in ("d_month", "d_category", "s_month", "s_category"):
            self.assertIn('name="%s"' % name, html, name)


class WorkshopCostEntryTests(SettlementFixtureMixin, TestCase):
    """车间费用页面：直接在表里填数据（不再上传），列顺序与用户模板一致，费用按车间顺序参与分配。"""

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="wc-tester", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.client.force_login(self.user)

    def _url(self):
        return reverse("workshop_costs", args=[self.period.id])

    def test_page_shows_template_columns_in_order(self):
        resp = self.client.get(self._url())
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        self.assertIn("<h1>车间费用</h1>", html)
        # 第一列是「期间」；列顺序＝期间｜顺序｜车间｜人工费｜…｜加工费｜合计｜备注
        heads = ["顺序", "期间", "车间", "人工费", "折旧费", "水电费", "蒸汽费",
                 "燃气费", "染料助剂费", "机物料费", "加工费", "合计", "备注"]
        positions = [html.index(">%s</th>" % name) for name in heads]
        self.assertEqual(positions, sorted(positions), heads)
        # 期间夹在「顺序」和「车间」中间
        self.assertLess(positions[0], positions[1])
        self.assertLess(positions[1], positions[2])
        # 六个车间都在（顺序＝结转顺序）
        for label in ("染色车间", "纺纱车间", "包覆车间", "针织车间", "机织车间", "浆粕车间"):
            self.assertIn(label, html, label)
        # 汇总表（按车间：每个车间一行，期间各一列）
        self.assertIn("车间费用汇总表（按车间）", html)
        self.assertIn(">车间</th>", html)
        self.assertNotIn("每个期间一行", html)

    def test_save_without_changes_does_not_reset(self):
        """只打开页面点保存、数字没改：不重置结算，也不把批次改成「页面录入」。"""
        uploaded = self.create_batch(self.period, ImportBatch.Kind.WORKSHOP_COST, "wc-up")
        pool = WorkshopCostPool.objects.create(
            period=self.period, batch=uploaded, workshop="RanZheng",
            wages=Decimal("100.00"), depreciation=Decimal("50.00"),
        )
        settlement = WorkshopSettlement.objects.create(
            period=self.period, workshop="RanZheng", sequence=1,
            status=WorkshopSettlement.Status.SETTLED,
        )
        data = {"csrfmiddlewaretoken": "x"}
        for field, _label in views.WORKSHOP_COST_COLUMNS:
            data["ws_%s_RanZheng_%s" % (self.period.id, field)] = "%.2f" % getattr(pool, field)
        resp = self.client.post(self._url(), data, follow=True)
        pool.refresh_from_db()
        settlement.refresh_from_db()
        self.assertEqual(pool.batch_id, uploaded.id)
        self.assertEqual(settlement.status, WorkshopSettlement.Status.SETTLED)
        self.assertIn("没有变化", resp.content.decode("utf-8"))

    def test_summary_table_is_cumulative_by_workshop(self):
        """汇总表列与录入表一致，数据是各期间累计（8月＝7月+8月）。"""
        august = AccountingPeriod.objects.create(month=date(2026, 8, 1))
        batch = self.create_batch(self.period, ImportBatch.Kind.WORKSHOP_COST, "wc-sum")
        WorkshopCostPool.objects.create(
            period=self.period, batch=batch, workshop="RanZheng",
            wages=Decimal("100.00"), depreciation=Decimal("50.00"),
        )
        WorkshopCostPool.objects.create(
            period=august, batch=batch, workshop="RanZheng",
            wages=Decimal("200.00"),
        )
        html = self.client.get(self._url()).content.decode("utf-8")
        block = html[html.index("车间费用汇总表（按车间）"):]
        # 列与「各车间费用」录入表一致
        for name in ("顺序", "期间", "车间", "人工费", "折旧费", "水电费", "蒸汽费",
                     "燃气费", "染料助剂费", "机物料费", "加工费", "合计", "备注"):
            self.assertIn(">%s</th>" % name, block, name)
        # 期间列显示累计区间
        self.assertIn("2026-07~2026-08", block)
        # 染色车间累计：人工 100+200=300、折旧 50 → 合计 350
        row_start = block.index("<strong>染色车间</strong>")
        row = block[row_start:block.index("</tr>", row_start)]
        self.assertIn("300.00", row)
        self.assertIn("50.00", row)
        self.assertIn("350.00", row)
        # 总合计行
        self.assertIn("总合计", block)
        self.assertIn("350.00", block)

    def test_entry_table_shows_only_current_period(self):
        """录入表只显示当前期间（7 月只填 7 月）；其它月份只出现在下面的汇总表里。"""
        august = AccountingPeriod.objects.create(month=date(2026, 8, 1))
        batch = self.create_batch(self.period, ImportBatch.Kind.WORKSHOP_COST, "wc-cur")
        WorkshopCostPool.objects.create(
            period=august, batch=batch, workshop="RanZheng", wages=Decimal("900.00"),
        )
        html = self.client.get(self._url()).content.decode("utf-8")
        self.assertIn("各车间费用（2026-07，单位：元）", html)
        entry = html[html.index('class="cost-entry"'):html.index("车间费用汇总表")]
        # 录入表里的输入框只属于当前期间（7 月）
        self.assertIn('name="ws_%s_RanZheng_wages"' % self.period.id, entry)
        self.assertNotIn('name="ws_%s_RanZheng_wages"' % august.id, entry)
        # 录入表里也只会出现当前期间的那一组行
        self.assertEqual(entry.count('data-period="%s"' % self.period.id), 6)
        self.assertNotIn('data-period="%s"' % august.id, entry)
        # 8 月的数据只在汇总表（累计）里体现
        summary = html[html.index("车间费用汇总表（按车间）"):]
        self.assertIn("2026-07~2026-08", summary)
        row_start = summary.index("<strong>染色车间</strong>")
        self.assertIn("900.00", summary[row_start:summary.index("</tr>", row_start)])

    def test_save_only_writes_current_period(self):
        """提交只写当前期间；别的期间在本页没有输入框，被塞进请求也不写库。"""
        august = AccountingPeriod.objects.create(month=date(2026, 8, 1))
        settlement = WorkshopSettlement.objects.create(
            period=self.period, workshop="RanZheng", sequence=1,
            status=WorkshopSettlement.Status.SETTLED,
        )
        batch = self.create_batch(self.period, ImportBatch.Kind.WORKSHOP_COST, "wc-half")
        WorkshopCostPool.objects.create(
            period=august, batch=batch, workshop="RanZheng", wages=Decimal("900.00"),
        )
        data = {"csrfmiddlewaretoken": "x"}
        for field, _label in views.WORKSHOP_COST_COLUMNS:
            data["ws_%s_RanZheng_%s" % (self.period.id, field)] = "0"
        data["ws_%s_RanZheng_wages" % self.period.id] = "800.00"
        data["ws_%s_RanZheng_wages" % august.id] = "123456.00"   # 故意塞进来，不应生效
        resp = self.client.post(self._url(), data, follow=True)
        self.assertEqual(resp.status_code, 200)
        self.assertEqual(
            WorkshopCostPool.objects.get(period=self.period, workshop="RanZheng").wages,
            Decimal("800.00"),
        )
        self.assertEqual(
            WorkshopCostPool.objects.get(period=august, workshop="RanZheng").wages,
            Decimal("900.00"),
        )
        settlement.refresh_from_db()
        self.assertEqual(settlement.status, WorkshopSettlement.Status.PENDING)
        self.assertIn('value="800.00"', resp.content.decode("utf-8"))

    def test_locked_period_cannot_be_saved(self):
        """已锁账期间提交 → 报错并跳回，不写库。"""
        self.period.status = AccountingPeriod.Status.LOCKED
        self.period.save(update_fields=["status"])
        data = {"csrfmiddlewaretoken": "x"}
        data["ws_%s_RanZheng_wages" % self.period.id] = "1.00"
        resp = self.client.post(self._url(), data, follow=True)
        self.assertEqual(resp.status_code, 200)
        self.assertIn("已锁账", resp.content.decode("utf-8"))
        self.assertFalse(
            WorkshopCostPool.objects.filter(period=self.period, workshop="RanZheng").exists()
        )

    def test_save_writes_pool_and_resets_settlement(self):
        settlement = WorkshopSettlement.objects.create(
            period=self.period,
            workshop="RanZheng",
            sequence=1,
            status=WorkshopSettlement.Status.SETTLED,
        )
        data = {"csrfmiddlewaretoken": "x"}
        for field, _label in views.WORKSHOP_COST_COLUMNS:
            data["ws_%s_RanZheng_%s" % (self.period.id, field)] = "100"
        data["ws_%s_RanZheng_wages" % self.period.id] = "1,234.56"   # 带千分位也要能存
        data["ws_%s_RanZheng_remark" % self.period.id] = "7月抄表"
        for code in ("tefang", "BaoFu", "ZhenZhi", "JiZhi", "浆粕车间"):
            for field, _label in views.WORKSHOP_COST_COLUMNS:
                data["ws_%s_%s_%s" % (self.period.id, code, field)] = "0"
        resp = self.client.post(self._url(), data, follow=True)
        self.assertEqual(resp.status_code, 200)
        pool = WorkshopCostPool.objects.get(period=self.period, workshop="RanZheng")
        self.assertEqual(pool.wages, Decimal("1234.56"))
        self.assertEqual(pool.water_power, Decimal("100.00"))
        self.assertEqual(pool.processing_fee, Decimal("100.00"))
        self.assertEqual(pool.remark, "7月抄表")
        # 保存后本期结算结果重置：车间结转回到待结转
        settlement.refresh_from_db()
        self.assertEqual(settlement.status, WorkshopSettlement.Status.PENDING)
        # 写进的是「页面录入」批次，和上传批次区分
        self.assertEqual(pool.batch.original_name, views.MANUAL_WORKSHOP_COST_NAME)
        # 页面回显新值
        html = resp.content.decode("utf-8")
        self.assertIn('value="1,234.56"', html)   # 录入值带千分位（用户 2026-09-21）
        self.assertIn("7月抄表", html)

    def test_invalid_number_is_rejected_and_value_kept(self):
        data = {"csrfmiddlewaretoken": "x"}
        for field, _label in views.WORKSHOP_COST_COLUMNS:
            data["ws_%s_RanZheng_%s" % (self.period.id, field)] = "0"
        data["ws_%s_RanZheng_wages" % self.period.id] = "abc"
        resp = self.client.post(self._url(), data)
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        self.assertIn("不是数字", html)
        # 没有落库
        self.assertFalse(
            WorkshopCostPool.objects.filter(period=self.period, workshop="RanZheng").exists()
        )
        # 用户填的内容还在输入框里
        self.assertIn('value="abc"', html)

    def test_negative_is_rejected(self):
        data = {"csrfmiddlewaretoken": "x"}
        for field, _label in views.WORKSHOP_COST_COLUMNS:
            data["ws_%s_RanZheng_%s" % (self.period.id, field)] = "0"
        data["ws_%s_RanZheng_steam" % self.period.id] = "-5"
        resp = self.client.post(self._url(), data)
        self.assertIn("不能为负数", resp.content.decode("utf-8"))

    def test_locked_period_is_read_only(self):
        """当前期间已锁账：提示只读、输入框禁用、不显示保存按钮。"""
        self.period.status = AccountingPeriod.Status.LOCKED
        self.period.save(update_fields=["status"])
        html = self.client.get(self._url()).content.decode("utf-8")
        self.assertIn("已锁账", html)
        self.assertIn("disabled", html)
        self.assertNotIn("保存车间费用", html)

    def test_nav_has_workshop_cost_under_cost_accounting(self):
        html = self.client.get(reverse("period_detail", args=[self.period.id])).content.decode("utf-8")
        self.assertIn(reverse("workshop_costs", args=[self.period.id]), html)
        # 位置：在「车间结转」之前
        self.assertLess(
            html.index(reverse("workshop_costs", args=[self.period.id])),
            html.index(reverse("workshop_settlements", args=[self.period.id])),
        )


class ProcessingFeeAllocationTests(SettlementFixtureMixin, TestCase):
    """加工费参与费用分配（不是每月都有：填了才分摊，为 0 时不影响任何结果）。"""

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="pf-tester", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.batch = self.create_batch(self.period, ImportBatch.Kind.INVENTORY, "pf")
        self.raw = Material.objects.create(code="RM-PF", name="加工费原料", uom="kg")
        # 同车间两张工单：各投料 100，材料成本各 1000（材料成本比例 = 各 50%）
        for order_no in ("WO-PF-1", "WO-PF-2"):
            WorkOrder.objects.create(
                period=self.period, order_no=order_no,
                production_dept_code="RanZheng", plan_quantity=Decimal("100"),
            )
        self.pool = WorkshopCostPool.objects.create(
            period=self.period,
            batch=self.create_batch(self.period, ImportBatch.Kind.WORKSHOP_COST, "pf-cost"),
            workshop="RanZheng",
            processing_fee=Decimal("1000.00"),
        )

    def _run(self):
        from settlement.services.workorder_costing import calculate_work_order_costs

        prices = {
            v.material_code: Decimal(v.weighted_average_price or 0)
            for v in self.period.inventory_valuations.all()
        }
        return calculate_work_order_costs(self.period, None, prices)

    def test_processing_fee_is_allocated_like_other_overhead(self):
        from settlement.services.workorder_costing import MANUFACTURING_FIELDS

        self.assertIn("processing_fee", MANUFACTURING_FIELDS)
        # 两张工单各领 100 公斤 × 10 元（用月末加权价兜底前先给价表）
        InventoryValuation.objects.create(
            period=self.period, material_code=self.raw.code,
            material_name=self.raw.name, weighted_average_price=Decimal("10"),
        )
        for order_no in ("WO-PF-1", "WO-PF-2"):
            WorkOrderDetail.objects.create(
                period=self.period, order_no=order_no,
                batch=self.batch,
                plan_quantity=Decimal("100"), completion_status="按时完成",
            )
            WorkReport.objects.create(
                period=self.period,
                batch=self.batch,
                report_time=datetime(2026, 7, 20, 8, 0),
                work_order_no=order_no,
                material_code=self.raw.code,
                material_name=self.raw.name,
                quantity=Decimal("100"),
            )
            MaterialIssue.objects.create(
                period=self.period,
                batch=self.batch,
                work_order_no=order_no,
                material_code=self.raw.code,
                material_name=self.raw.name,
                material_category="原料",
                issue_quantity=Decimal("100"),
                returned_quantity=Decimal("0"),
                actual_issue_quantity=Decimal("100"),
                issue_time=datetime(2026, 7, 10, 8, 0),
            )
        self._run()
        results = list(self.period.work_order_cost_results.all())
        self.assertEqual(len(results), 2)
        total_fee = sum(r.processing_fee for r in results)
        # 加工费 1000 全额分摊到两张工单（材料成本比例各 50%）
        self.assertEqual(total_fee, Decimal("1000.00"))
        for r in results:
            self.assertEqual(r.processing_fee, Decimal("500.00"))
            # 加工费属于制造费用：计入完工成本（本用例全部完工）
            self.assertEqual(r.finished_cost, r.finished_material + r.finished_overhead + r.finished_wages)
            self.assertIn(r.processing_fee, (r.finished_overhead,))
        # 分摊金额已进入工单总成本
        self.assertGreater(results[0].total_cost, results[0].direct_material)

    def test_page_and_export_show_processing_fee_column(self):
        """工单成本页与导出都带「加工费 / 加工费分配率」两列。"""
        self.client.force_login(self.user)
        work_order = WorkOrder.objects.get(period=self.period, order_no="WO-PF-1")
        WorkOrderCostResult.objects.create(
            period=self.period, work_order=work_order,
            product_code="P-PF", product_name="加工费试算",
            plan_quantity=Decimal("100"), finished_quantity=Decimal("100"),
            direct_material=Decimal("1000.00"),
            processing_fee=Decimal("250.00"),
            total_cost=Decimal("1250.00"), finished_cost=Decimal("1250.00"),
            finished_material=Decimal("1000.00"),
            finished_overhead=Decimal("250.00"),
            finished_unit_cost=Decimal("12.5"),
        )
        html = self.client.get(
            reverse("work_order_costs", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertIn("加工费", html)
        self.assertIn("加工费分配率", html)
        self.assertIn("250.00", html)

    def test_zero_processing_fee_keeps_old_numbers(self):
        """没填加工费的月份：加工费为 0，不影响原有费用分摊结果。"""
        InventoryValuation.objects.create(
            period=self.period, material_code=self.raw.code,
            material_name=self.raw.name, weighted_average_price=Decimal("10"),
        )
        WorkOrderDetail.objects.create(
            period=self.period, order_no="WO-PF-1",
            batch=self.batch,
            plan_quantity=Decimal("100"), completion_status="按时完成",
        )
        WorkReport.objects.create(
            period=self.period, batch=self.batch,
            report_time=datetime(2026, 7, 20, 8, 0), work_order_no="WO-PF-1",
            material_code=self.raw.code, material_name=self.raw.name,
            quantity=Decimal("100"),
        )
        MaterialIssue.objects.create(
            period=self.period, batch=self.batch, work_order_no="WO-PF-1",
            material_code=self.raw.code, material_name=self.raw.name,
            material_category="原料", issue_quantity=Decimal("100"),
            returned_quantity=Decimal("0"), actual_issue_quantity=Decimal("100"),
            issue_time=datetime(2026, 7, 10, 8, 0),
        )
        self.pool.processing_fee = Decimal("0.00")
        self.pool.save(update_fields=["processing_fee"])
        self._run()
        result = self.period.work_order_cost_results.get(work_order__order_no="WO-PF-1")
        self.assertEqual(result.processing_fee, Decimal("0.00"))
        self.assertEqual(result.water_power, Decimal("0.00"))
        self.assertEqual(result.total_cost, result.finished_cost + result.wip_cost)


class FinishedReceiptVoucherTests(SettlementFixtureMixin, TestCase):
    """完工入库结转凭证（用户 2026-09-25；贷方改成本要素 2026-09-27）：

    - 摘要按车间：{月}月{车间}完工入库{物料分类}（贷方要素行 {月}月各车间完工入库{要素}）；
    - 科目按物料分类：借 库存商品—物料分类；贷 生产成本—直接材料/直接人工/制造费用（各车间汇总）；
    - 金额基准＝工单完工成本（报工数量 × 完工单位成本），贷方要素取完工直接材料/人工/制造费用；
    - 浆粕车间没有工单，按库存流水「生产入库」单独归集（贷方一行 生产成本—浆粕车间，不拆要素）；
    - 卡片上不再有借贷金额 pill 与附件 1/2，明细统一到「打开明细」页。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="receipt-voucher", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.batch = self.create_batch(
            self.period, ImportBatch.Kind.WORK_ORDER, "receipt-voucher"
        )
        # 染色车间一张工单：完工 10、完工成本 300（产品编码能查到分类=染色纤维）
        self.order = WorkOrder.objects.create(
            period=self.period, batch=self.batch, order_no="MO-RC-1",
            production_dept_code="RanZheng", plan_quantity=Decimal("10"),
        )
        WorkOrderCostResult.objects.create(
            period=self.period, work_order=self.order, product_code="R51.202",
            product_name="染色纤维", plan_quantity=Decimal("10"),
            finished_quantity=Decimal("10"), wip_quantity=Decimal("0"),
            direct_material=Decimal("200"), total_cost=Decimal("300"),
            finished_cost=Decimal("300"), wip_cost=Decimal("0"),
            finished_unit_cost=Decimal("30"),
        )
        # 浆粕车间：没有工单，一个成品「生产入库」流水 50 元
        Material.objects.create(code="21.03", name="芳纶浆粕DFP-103", uom="kg")
        pulp_batch = self.create_batch(
            self.period, ImportBatch.Kind.PULP_FLOW, "receipt-voucher-pulp"
        )
        txn = self.create_txn(
            self.period, pulp_batch, txn_date=date(2026, 7, 20),
            document_no="", business_type="生产入库",
            direction=InventoryTransaction.Direction.IN,
            material=Material.objects.get(code="21.03"),
            quantity=Decimal("5"),
        )
        txn.amount = Decimal("50.00")
        txn.category_name = "浆粕"
        txn.save(update_fields=["amount", "category_name"])

    def test_summary_by_workshop_account_by_category(self):
        data = build_finished_receipt_voucher(self.period)
        line_rows = [row for row in data["rows"] if row["level"] == "line"]
        debit_rows = [row for row in line_rows if row["debit"]]
        credit_rows = [row for row in line_rows if row["credit"]]
        # 借方：库存商品—物料分类，摘要带车间
        self.assertTrue(any(
            row["account"] == "库存商品—浆粕" for row in debit_rows
        ))
        for row in debit_rows:
            self.assertTrue(row["account"].startswith("库存商品"), row["account"])
            self.assertIn("完工入库", row["summary"])
        # 贷方：生产成本—直接材料/直接人工/制造费用（全车间汇总，含浆粕，各一行）
        # （用户 2026-09-27 改：贷方二级科目＝成本要素；浆粕也并进这三个要素）
        self.assertEqual(
            {row["account"] for row in credit_rows},
            {"生产成本—直接材料"},
        )
        by_account = {row["account"]: row for row in credit_rows}
        # fixture 的染色工单只填了 finished_cost=300（没拆要素）→ 差额并入直接材料；
        # 浆粕 50 元没传浆粕原料流水、拆分取不到 → 也并入直接材料
        self.assertEqual(by_account["生产成本—直接材料"]["credit"], Decimal("350.00"))
        self.assertEqual(
            by_account["生产成本—直接材料"]["summary"], "7月各车间完工入库直接材料"
        )
        # 要素合计＝工单完工成本（不含浆粕）
        self.assertEqual(
            data["summary"]["element_totals"],
            {"直接材料": Decimal("350.00"), "直接人工": Decimal("0.00"),
             "制造费用": Decimal("0.00")},
        )
        # 拆分取不到时有提示
        self.assertTrue(any("全额并入直接材料" in c["text"] for c in data["checks"]))
        # 借贷平衡：300（染色）+ 50（浆粕）= 350
        self.assertTrue(data["totals"]["balanced"])
        self.assertEqual(data["totals"]["debit"], Decimal("350.00"))
        self.assertEqual(data["totals"]["credit"], Decimal("350.00"))
        self.assertEqual(
            sum((row["amount"] for row in data["category_rows"]), Decimal("0")),
            Decimal("350.00"),
        )

    def test_summary_text_uses_workshop_label(self):
        data = build_finished_receipt_voucher(self.period)
        summaries = {row["summary"] for row in data["rows"] if row["debit"]}
        self.assertIn("7月染色车间完工入库染色纤维", summaries)
        self.assertIn("7月浆粕车间完工入库浆粕", summaries)
        # 贷方要素行的摘要（各车间汇总）
        credit_summaries = {row["summary"] for row in data["rows"] if row["credit"]}
        self.assertIn("7月各车间完工入库直接材料", credit_summaries)

    def test_credit_splits_into_cost_elements(self):
        """贷方按成本要素分行（用户 2026-09-27）：直接材料 / 直接人工 / 制造费用。

        金额取工单成本的完工直接材料 / 完工人工 / 完工制造费用（三者相加＝完工成本），
        仍然不用库存流水；浆粕车间并进这三个要素（按浆粕车间成本归集拆）。
        """
        order = WorkOrder.objects.create(
            period=self.period, batch=self.batch, order_no="MO-RC-ELEM",
            production_dept_code="tefang", plan_quantity=Decimal("10"),
        )
        WorkOrderCostResult.objects.create(
            period=self.period, work_order=order, product_code="11.01",
            product_name="纱线", plan_quantity=Decimal("10"),
            finished_quantity=Decimal("10"), wip_quantity=Decimal("0"),
            direct_material=Decimal("300"), wages=Decimal("80"),
            finished_material=Decimal("300"), finished_wages=Decimal("80"),
            finished_overhead=Decimal("120"), finished_cost=Decimal("500"),
            total_cost=Decimal("500"), wip_cost=Decimal("0"),
            finished_unit_cost=Decimal("50"),
        )
        data = build_finished_receipt_voucher(self.period)
        credit = {
            row["account"]: row["credit"]
            for row in data["rows"] if row["credit"] and row["level"] == "line"
        }
        # 染色工单 300 未拆要素 → 直接材料 300；本单 材料 300 / 人工 80 / 制造费用 120；
        # 浆粕 50（拆分取不到）→ 并入直接材料
        self.assertEqual(credit["生产成本—直接材料"], Decimal("650.00"))
        self.assertEqual(credit["生产成本—直接人工"], Decimal("80.00"))
        self.assertEqual(credit["生产成本—制造费用"], Decimal("120.00"))
        self.assertEqual(set(credit), {
            "生产成本—直接材料", "生产成本—直接人工", "生产成本—制造费用",
        })
        self.assertEqual(
            data["summary"]["element_totals"],
            {"直接材料": Decimal("650.00"), "直接人工": Decimal("80.00"),
             "制造费用": Decimal("120.00")},
        )
        # 借贷平衡：800（工单完工成本）+ 50（浆粕）= 850
        self.assertTrue(data["totals"]["balanced"])
        self.assertEqual(data["totals"]["debit"], Decimal("850.00"))
        self.assertEqual(data["totals"]["credit"], Decimal("850.00"))
        # 附件3：按车间 × 成本要素（浆粕不拆要素）
        by_workshop = {row["workshop"]: row for row in data["workshop_rows"]}
        self.assertTrue(by_workshop["tefang"]["element_split"])
        self.assertEqual(by_workshop["tefang"]["material"], Decimal("300.00"))
        self.assertEqual(by_workshop["tefang"]["wages"], Decimal("80.00"))
        self.assertEqual(by_workshop["tefang"]["overhead"], Decimal("120.00"))
        self.assertEqual(by_workshop["RanZheng"]["material"], Decimal("300.00"))
        # 浆粕：归集取不到 → 全额挂直接材料，element_split=False
        self.assertFalse(by_workshop["浆粕车间"]["element_split"])
        self.assertEqual(by_workshop["浆粕车间"]["material"], Decimal("50.00"))
        self.assertEqual(by_workshop["浆粕车间"]["amount"], Decimal("50.00"))
        self.assertEqual(data["workshop_totals"]["amount"], Decimal("850.00"))

    def test_pulp_also_splits_into_elements(self):
        """浆粕车间的贷方也拆直接材料/直接人工/制造费用（用户 2026-09-27）。

        金额仍取库存流水「生产入库」；拆分取浆粕车间成本归集：
        直接材料＝浆粕原料领料、直接人工＝车间工资、制造费用＝其余车间费用。
        """
        from settlement.models import PulpMaterial, WorkshopCostPool

        pulp_batch = self.create_batch(
            self.period, ImportBatch.Kind.PULP_FLOW, "receipt-voucher-pulp-split"
        )
        # 生产入库金额＝浆粕完工成本（材料 300 + 工资 100 + 制造费用 100 = 500）
        txn = self.period.inventory_transactions.filter(
            business_type="生产入库", material_code="21.03"
        ).first()
        txn.amount = Decimal("500.00")
        txn.unit_price = Decimal("100")
        txn.save(update_fields=["amount", "unit_price"])
        PulpMaterial.objects.create(
            period=self.period, batch=pulp_batch, txn_date=date(2026, 7, 5),
            document_no="PULP-IN", material_code="R52.101",
            material_name="对位芳纶短纤维", quantity=Decimal("8"),
            unit_price=Decimal("37.5"), amount=Decimal("300"),
            business_type="领料出库", uom="kg",
        )
        WorkshopCostPool.objects.create(
            period=self.period, batch=pulp_batch, workshop="浆粕车间",
            water_power=Decimal("100"), wages=Decimal("100"),
        )
        data = build_finished_receipt_voucher(self.period)
        credit = {
            row["account"]: row["credit"]
            for row in data["rows"] if row["credit"] and row["level"] == "line"
        }
        # 染色工单 300（未拆要素→直接材料）+ 浆粕：材料 300 / 人工 100 / 制造费用 100
        self.assertEqual(credit["生产成本—直接材料"], Decimal("600.00"))
        self.assertEqual(credit["生产成本—直接人工"], Decimal("100.00"))
        self.assertEqual(credit["生产成本—制造费用"], Decimal("100.00"))
        # 浆粕单独的部分也能拆出来
        self.assertEqual(
            data["summary"]["pulp_element_totals"],
            {"直接材料": Decimal("300.00"), "直接人工": Decimal("100.00"),
             "制造费用": Decimal("100.00")},
        )
        self.assertTrue(data["summary"]["pulp_split_ok"])
        self.assertEqual(data["summary"]["pulp_amount"], Decimal("500.00"))
        # 借贷平衡：300（染色）+ 500（浆粕）= 800
        self.assertTrue(data["totals"]["balanced"])
        self.assertEqual(data["totals"]["debit"], Decimal("800.00"))
        # 附件3：浆粕行有要素拆分（不再是 —）
        pulp_row = next(r for r in data["workshop_rows"] if r["workshop"] == "浆粕车间")
        self.assertTrue(pulp_row["element_split"])
        self.assertEqual(pulp_row["material"], Decimal("300.00"))
        self.assertEqual(pulp_row["wages"], Decimal("100.00"))
        self.assertEqual(pulp_row["overhead"], Decimal("100.00"))
        # 附件1 的浆粕逐单行也带要素
        self.assertEqual(data["pulp_rows"][0]["material"], Decimal("300.00"))

    def test_detail_page_renders(self):
        """「打开明细」页：贷方要素 + 三张附件（原卡片上的附件1/2 搬到这里）。"""
        self.client.force_login(self.user)
        resp = self.client.get(
            reverse("finished_receipt_voucher", args=[self.period.id])
        )
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        self.assertIn("完工入库结转凭证", html)
        self.assertIn("生产成本—直接材料", html)
        self.assertIn("生产成本—直接人工", html)
        self.assertIn("生产成本—制造费用", html)
        # 浆粕车间也并进这三个要素（用户 2026-09-27）
        self.assertNotIn("生产成本—浆粕车间", html)
        self.assertIn("附件1：逐单完工入库明细", html)
        self.assertIn("附件2：按车间 × 物料分类汇总", html)
        self.assertIn("附件3：按车间 × 成本要素", html)

    def test_pulp_is_collected_separately(self):
        data = build_finished_receipt_voucher(self.period)
        self.assertEqual(data["summary"]["pulp_count"], 1)
        self.assertEqual(data["summary"]["pulp_amount"], Decimal("50.00"))
        row = data["pulp_rows"][0]
        self.assertEqual(row["source"], "库存流水·生产入库")
        self.assertEqual(row["category"], "浆粕")
        # 工单侧只有 1 张（染色），不含浆粕
        self.assertEqual(data["summary"]["order_count"], 1)
        self.assertEqual(
            {r["order_no"] for r in data["detail_rows"]}, {"MO-RC-1"}
        )

    def test_duplicate_receipt_flow_is_flagged(self):
        """报工入库流水有重复行 → blocking 提示，且凭证按工单完工成本计（不重复）。"""
        order = WorkOrder.objects.create(
            period=self.period, batch=self.batch, order_no="MO-RC-DUP",
            production_dept_code="tefang", plan_quantity=Decimal("5"),
        )
        WorkOrderCostResult.objects.create(
            period=self.period, work_order=order, product_code="11.01",
            product_name="纱线", plan_quantity=Decimal("5"),
            finished_quantity=Decimal("5"), wip_quantity=Decimal("0"),
            direct_material=Decimal("100"), total_cost=Decimal("100"),
            finished_cost=Decimal("100"), wip_cost=Decimal("0"),
            finished_unit_cost=Decimal("20"),
        )
        material = Material.objects.create(code="11.01", name="对位芳纶纱线", uom="kg")
        for _ in range(2):   # 同一单据、同一天、同数量、同金额的两行 → 重复行
            txn = self.create_txn(
                self.period, self.batch, txn_date=date(2026, 7, 10),
                document_no="MO-RC-DUP", business_type="报工",
                direction=InventoryTransaction.Direction.IN,
                material=material, quantity=Decimal("5"),
            )
            txn.amount = Decimal("100.00")
            txn.category_name = "纱线"
            txn.save(update_fields=["amount", "category_name"])
        data = build_finished_receipt_voucher(self.period)
        self.assertEqual(data["summary"]["duplicate_rows"], 1)
        self.assertEqual(data["summary"]["duplicate_amount"], Decimal("100.00"))
        self.assertTrue(any(
            "重复行" in check["text"] for check in data["checks"]
        ))
        # 凭证里这张单只计一次（工单完工成本 100），不是 200
        row = next(r for r in data["detail_rows"] if r["order_no"] == "MO-RC-DUP")
        self.assertEqual(row["amount"], Decimal("100.00"))
        self.assertEqual(data["totals"]["debit"], Decimal("450.00"))

    def test_export_has_three_sheets(self):
        self.client.force_login(self.user)
        resp = self.client.get(
            reverse("export_listing", args=[self.period.id, "finished_receipt_voucher"])
        )
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(BytesIO(resp.content))
        self.assertEqual(
            wb.sheetnames,
            ["凭证分录", "附件1-逐单完工入库", "附件2-车间×物料分类",
             "附件3-车间×成本要素"],
        )
        sheet = wb["凭证分录"]
        header = [cell.value for cell in sheet[1]]
        self.assertEqual(
            header,
            ["凭证期间", "摘要", "科目", "辅助说明", "数量", "借方", "贷方"],
        )
        # 贷方按成本要素：生产成本—直接材料（fixture 人工/制造费用为 0 → 该行不出现）
        # 浆粕不再单独挂「生产成本—浆粕车间」，而是并进这三个要素
        accounts = [row[2].value for row in sheet.iter_rows(min_row=2)]
        self.assertIn("生产成本—直接材料", accounts)
        self.assertNotIn("生产成本—浆粕车间", accounts)
        # 附件1 带三个要素列
        detail_header = [cell.value for cell in wb["附件1-逐单完工入库"][1]]
        self.assertIn("直接材料", detail_header)
        self.assertIn("直接人工", detail_header)
        self.assertIn("制造费用", detail_header)
        # 末行合计：借＝贷
        last = [cell.value for cell in sheet[sheet.max_row]]
        self.assertEqual(Decimal(str(last[5])), Decimal(str(last[6])))
        self.assertEqual(Decimal(str(last[5])), Decimal("350.00"))

    def test_voucher_center_has_the_card(self):
        self.client.force_login(self.user)
        html = self.client.get(
            reverse("voucher_center", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertIn("完工入库结转凭证", html)
        self.assertIn("库存商品—浆粕", html)
        self.assertIn("生产成本—直接材料", html)
        # 浆粕车间不再单独挂「生产成本—浆粕车间」，而是并进上面三个要素
        self.assertNotIn("生产成本—浆粕车间", html)
        # 用户 2026-09-27：卡片不再带附件1/2（搬到「打开明细」页），也不再显示借贷金额 pill
        self.assertNotIn("附件1-按车间", html)
        self.assertNotIn("附件2-按物料分类（借方科目）", html)
        self.assertIn("打开明细", html)
        self.assertIn(
            reverse("finished_receipt_voucher", args=[self.period.id]), html
        )
        card = next(
            item for item in build_voucher_center(self.period)["cards"]
            if item["key"] == "production_receipt"
        )
        self.assertIsNone(card["balance"])
        self.assertNotIn("extra_tables", card)
        self.assertEqual(card["url_name"], "finished_receipt_voucher")


class VoucherCenterTests(SettlementFixtureMixin, TestCase):
    """凭证中心：一页列出全部凭证（已实现 2 类 + 规划中 3 类）。"""

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="vc-tester", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.batch = self.create_batch(self.period, ImportBatch.Kind.INVENTORY, "vc")
        self.material = Material.objects.create(code="RM-VC", name="凭证中心原料", uom="kg")

    def _sales_line(self, row_no=1, *, carry_type="当月发货当月开票结转成本",
                    category="纱线", quantity="10", cost="100.00",
                    out_month=None, invoice_month=None, status="invoiced"):
        batch = self.create_batch(self.period, ImportBatch.Kind.SALES_LEDGER,
                                  "vc-sales-%s" % row_no)
        return SalesLedgerLine.objects.create(
            period=self.period,
            batch=batch,
            row_no=row_no,
            out_month=out_month or self.period.month,
            invoice_status=status,
            invoice_month=invoice_month or self.period.month,
            channel="国内销售",
            order_no="SO-VC-%s" % row_no,
            customer_code="C1",
            customer_name="客户甲",
            sales_category=category,
            material_code="11.01",
            material_name="产品",
            uom="kg",
            quantity=Decimal(quantity),
            terminal_amount=Decimal("200.00"),
            cost_amount=Decimal(cost),
            carry_type=carry_type,
        )

    def _receipt(self, quantity=Decimal("100"), unit_price=Decimal("10")):
        txn = self.create_txn(
            self.period,
            self.batch,
            txn_date=date(2026, 7, 5),
            document_no="GR-VC",
            business_type="采购入库",
            direction=InventoryTransaction.Direction.IN,
            material=self.material,
            quantity=quantity,
            supplier_code="SUP-VC",
            supplier_name="凭证中心供应商",
        )
        txn.unit_price = unit_price
        txn.amount = (quantity * unit_price).quantize(Decimal("0.01"))
        txn.price_source = "手工暂估"
        txn.category_name = "原料"
        txn.save()
        return txn

    def test_center_lists_ready_and_planned_vouchers(self):
        self._receipt()
        data = build_voucher_center(self.period)
        # 用户 2026-09-25：完工入库结转凭证已实现 → 已实现 4 类、规划中 1 类
        self.assertEqual(data["ready_count"], 4)
        self.assertEqual(data["planned_count"], 1)
        # 卡片顺序＝建议的生成顺序：暂估回冲 → 投料 → 制造费用 → 完工入库 → 销售结转
        keys = [card["key"] for card in data["cards"]]
        self.assertEqual(
            keys,
            ["purchase", "material_issue", "manufacturing_overhead",
             "production_receipt", "sales"],
        )
        self.assertEqual([card["step"] for card in data["cards"]], [1, 2, 3, 4, 5])
        # 末尾「生成顺序」列表与卡片顺序一致
        self.assertEqual([step["step"] for step in data["steps"]], [1, 2, 3, 4, 5])
        self.assertEqual(
            [step["key"] for step in data["steps"]], keys
        )
        purchase = {card["key"]: card for card in data["cards"]}["purchase"]
        # 与销售卡片一致：不展示汇总字段与借贷 pill
        self.assertIsNone(purchase["balance"])
        self.assertEqual(purchase["metrics"], [])
        self.assertTrue(purchase["entries"])
        self.assertTrue(data["all_balanced"])

    def test_center_aggregates_blocking_checks(self):
        material = Material.objects.create(code="RM-VC2", name="无价原料", uom="kg")
        self.create_txn(
            self.period,
            self.batch,
            txn_date=date(2026, 7, 6),
            document_no="GR-VC2",
            business_type="采购入库",
            direction=InventoryTransaction.Direction.IN,
            material=material,
            quantity=Decimal("5"),
            supplier_code="SUP-VC",
            supplier_name="凭证中心供应商",
        )
        data = build_voucher_center(self.period)
        self.assertTrue(data["blocking"])
        self.assertTrue(
            any("没有暂估单价" in check["text"] for check in data["blocking"])
        )

    def test_center_page_renders_every_voucher(self):
        self._receipt()
        self.client.force_login(self.user)
        url = reverse("voucher_center", args=[self.period.id])
        html = self.client.get(url).content.decode("utf-8")
        for text in (
            "<h1>凭证中心</h1>",
            "销售结转成本凭证",
            "采购暂估 / 回冲凭证",
            "生产领料（投料）凭证",
            "制造费用 / 人工结转凭证",
            "完工入库结转凭证",
            "建议的凭证生成顺序",
        ):
            self.assertIn(text, html, text)
        # 已实现的凭证给出打开明细与导出入口
        self.assertIn(reverse("sales_cost_voucher", args=[self.period.id]), html)
        self.assertIn(reverse("purchase_accrual_voucher", args=[self.period.id]), html)
        self.assertIn(
            reverse("export_listing", args=[self.period.id, "purchase_voucher"]), html
        )
        # 卡片按生成顺序编号 + 锚点，末尾顺序表可跳转
        for step, key in enumerate(
            ["purchase", "material_issue", "manufacturing_overhead",
             "production_receipt", "sales"], start=1
        ):
            self.assertIn('id="voucher-%s"' % key, html, key)
            self.assertIn(">第 %d 步<" % step, html, key)
            self.assertIn('href="#voucher-%s"' % key, html, key)
        # 卡片出现顺序＝编号顺序
        positions = [html.index('id="voucher-%s"' % key) for key in
                     ["purchase", "material_issue", "manufacturing_overhead",
                      "production_receipt", "sales"]]
        self.assertEqual(positions, sorted(positions))


    def test_center_tables_match_voucher_detail_format(self):
        """凭证中心里的表格式必须与凭证页一致：列名、列序、列数都不能走样。"""
        self._receipt()
        data = build_voucher_center(self.period)
        cards = {card["key"]: card for card in data["cards"]}
        self.assertEqual(
            cards["sales"]["entry_headers"], ["摘要", "科目", "数量", "借方", "贷方"]
        )
        self.assertEqual(
            cards["purchase"]["entry_headers"], ["摘要", "科目", "数量", "借方", "贷方"]
        )
        self.assertEqual(cards["purchase"]["metrics"], [])
        self.assertIsNone(cards["purchase"]["balance"])
        for key in ("sales", "purchase"):
            width = len(cards[key]["entry_headers"])
            for row in cards[key]["entries"]:
                self.assertEqual(len(row["cells"]), width, key)

    def test_center_sales_table_headers_match_detail_page(self):
        """页面渲染层面：中心里那张销售凭证表，表头与明细页逐列一致。"""
        self._receipt()
        self._sales_line()
        self.client.force_login(self.user)
        center = self.client.get(
            reverse("voucher_center", args=[self.period.id])
        ).content.decode("utf-8")
        detail = self.client.get(
            reverse("sales_cost_voucher", args=[self.period.id])
        ).content.decode("utf-8")

        sales_idx = center.index('id="voucher-sales"')
        center_body = center[sales_idx:center.index("</table>", sales_idx)]
        center_heads = [
            re.sub(r"<[^>]+>", "", h).strip()
            for h in re.findall(r"<th[^>]*>(.*?)</th>", center_body, re.S)
        ]

        detail_idx = detail.index("<h2>凭证分录（不含税）</h2>")
        detail_body = detail[detail_idx:detail.index("</table>", detail_idx)]
        detail_heads = [
            re.sub(r"<[^>]+>", "", h).strip()
            for h in re.findall(r"<th[^>]*>(.*?)</th>", detail_body, re.S)
        ]

        self.assertEqual(detail_heads, ["摘要", "科目", "数量", "借方", "贷方"])
        self.assertEqual(center_heads, detail_heads)

        # 首行第一列是会计科目（不是结转类型）
        tbody = center_body[center_body.index("<tbody>"):]
        first_tr = tbody[:tbody.index("</tr>")]
        cells = [
            re.sub(r"<[^>]+>", "", c).strip()
            for c in re.findall(r"<td[^>]*>(.*?)</td>", first_tr, re.S)
        ]
        # 中心与明细页同口径：摘要＝结转类型 + 大类，科目借方到渠道、贷方带销售大类
        self.assertTrue(cells[0].startswith("当月发货当月开票结转成本"), cells[0])
        self.assertTrue(cells[1].startswith("主营业务成本—"), cells[1])

    def test_sales_voucher_uses_entry_format_like_purchase(self):
        """销售凭证与采购凭证同格式：汇总卡 + 生成前检查 + 凭证分录（借/贷）+ 附件。"""
        self._receipt()
        self.client.force_login(self.user)
        html = self.client.get(
            reverse("sales_cost_voucher", args=[self.period.id])
        ).content.decode("utf-8")
        for text in ("<h1>销售结转成本凭证</h1>",
                     "生成前检查", "<h2>凭证分录（不含税）</h2>",
                     "<h2>附件：结转类型 × 销售大类（含表内筛选）</h2>",
                     "取数口径与说明", "voucher-filter-form", "凭证期间"):
            self.assertIn(text, html, text)
        # 汇总卡与借贷 pill 按要求去掉
        self.assertNotIn("summary-strip", html)
        self.assertNotIn("参与结转台账行数", html)
        # 采购凭证页也是同一套结构，保证两页格式一致
        purchase = self.client.get(
            reverse("purchase_accrual_voucher", args=[self.period.id])
        ).content.decode("utf-8")
        for text in ("生成前检查", "<h2>凭证分录（不含税）</h2>", "凭证期间"):
            self.assertIn(text, purchase, text)
        self.assertNotIn("summary-strip", purchase)

    def test_nav_has_single_voucher_center_entry(self):
        """导航合并成一个「凭证中心」入口，明细页保留返回入口。"""
        self.client.force_login(self.user)
        center = reverse("voucher_center", args=[self.period.id])
        html = self.client.get(center).content.decode("utf-8")
        self.assertIn("凭证中心", html)
        self.assertNotIn(
            ">结转成本凭证</a>", html
        )
        for name in ("sales_cost_voucher", "purchase_accrual_voucher"):
            detail = self.client.get(
                reverse(name, args=[self.period.id])
            ).content.decode("utf-8")
            self.assertIn(center, detail, name)


class KeepScrollPositionTests(TestCase):
    """筛选/切换条件提交后不跳回页面顶部（全站生效）。"""

    def setUp(self):
        self.base_tpl = (settings.BASE_DIR / "templates" / "base.html").read_text(
            encoding="utf-8"
        )

    def test_base_saves_and_restores_scroll(self):
        for text in ("keepScrollPos", "sessionStorage.setItem", "window.scrollY",
                     "window.scrollTo"):
            self.assertIn(text, self.base_tpl, text)

    def test_listens_in_capture_phase(self):
        # 必须用捕获阶段：要早于 onchange="this.form.submit()" 这类内联处理器
        self.assertIn('document.addEventListener("change"', self.base_tpl)
        self.assertIn('document.addEventListener("click"', self.base_tpl)
        self.assertIn("}, true);", self.base_tpl)

    def test_only_get_forms(self):
        # 只处理筛选用 GET 表单（POST 表单如保存/锁账不受影响）
        self.assertIn("function isGetForm(form)", self.base_tpl)
        self.assertIn('(form.getAttribute("method") || "get").toLowerCase() === "get"',
                      self.base_tpl)

    def test_has_max_age_guard(self):
        # 过期的记录不恢复（避免很久以后再打开页面突然跳到某个位置）
        self.assertIn("MAX_AGE_MS", self.base_tpl)
        self.assertIn("Date.now() - (data.t || 0) < MAX_AGE_MS", self.base_tpl)

    def test_restores_on_load(self):
        self.assertIn('window.addEventListener("load"', self.base_tpl)


class SalesLedgerPageTests(SettlementFixtureMixin, TestCase):
    """销售台账页：逐行明细 + 筛选 + 合计 + 导出（与页面同一套筛选）。"""

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="sl-page", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.batch = self.create_batch(
            self.period, ImportBatch.Kind.SALES_LEDGER, "slp"
        )
        self.url = reverse("sales_ledger", args=[self.period.id])
        self.june = date(2026, 6, 1)

    _UNSET = object()

    def _line(self, row_no, *, out_month=None, status="invoiced", quantity="10",
              revenue="1000", cost_price=None, cost=None, customer="张三",
              category="纱线", channel="国内销售", margin=None,
              profit=None, invoice_month=_UNSET):
        return SalesLedgerLine.objects.create(
            period=self.period,
            batch=self.batch,
            row_no=row_no,
            out_month=out_month or self.period.month,
            invoice_status=status,
            invoice_month=(
                self.period.month
                if invoice_month is self._UNSET and status == "invoiced"
                else (None if invoice_month is self._UNSET else invoice_month)
            ),
            channel=channel,
            customer_code="C001",
            customer_name=customer,
            sales_category=category,
            material_code="21.03",
            material_name="21.03 芳纶浆粕",
            uom="kg",
            terminal_amount=Decimal(revenue) if revenue is not None else None,
            quantity=Decimal(quantity),
            cost_price=Decimal(cost_price) if cost_price is not None else None,
            cost_amount=Decimal(cost) if cost is not None else None,
            gross_profit=Decimal(profit) if profit is not None else None,
            gross_margin=Decimal(margin) if margin is not None else None,
        )

    # ———————————— 页面 ————————————

    def test_page_renders_with_all_columns(self):
        self._line(1, cost_price="80", cost="800", profit="200", margin="0.2")
        resp = self.client.get(self.url)
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        for header in ("销售台账", "出库期间", "开票状态", "销售数量",
                       "终端销售收入", "成本单价", "销售成本",
                       "销售毛利", "毛利率"):
            self.assertIn(header, html, header)

    def test_totals_match_queryset(self):
        self._line(1, cost_price="80", cost="800", profit="200", margin="0.2")
        self._line(2, quantity="5", revenue="500", cost_price="80", cost="400",
                   profit="100", margin="0.2")
        resp = self.client.get(self.url)
        html = resp.content.decode("utf-8")
        # 合计金额（千分位）与页面上一致（汇总表合计行 + 明细合计行）
        self.assertIn("1,500.00", html)   # 终端收入合计
        self.assertIn("1,200.00", html)   # 销售成本合计
        self.assertIn("300.00", html)     # 销售毛利合计

    def test_summary_is_table_grouped_by_period_and_category(self):
        # 用户要求：原来那排卡片（明细行数…毛利率）改成表格，
        # 字段 = 期间 / 销售大类 / 终端销售收入 / 销售成本 / 销售毛利 / 毛利率
        self._line(1, category="纱线", cost_price="80", cost="800", profit="200", margin="0.2")
        self._line(2, category="面料", quantity="5", revenue="500",
                   cost_price="80", cost="400", profit="100", margin="0.2")
        # 第三行：出库与开票都在 6 月 -> 汇总归到 2026-06 组
        self._line(3, out_month=self.june, invoice_month=self.june, category="纱线",
                   quantity="2", revenue="200", cost_price="50", cost="100",
                   profit="100", margin="0.5")
        # 默认只看当月，这里显式取「全部」来验证分期分组
        html = self.client.get(self.url, {"invoice_month": ""}).content.decode("utf-8")
        # 不再是卡片排，改成表格
        self.assertNotIn('class="summary-strip"', html)
        self.assertIn('class="sales-summary-table"', html)
        block = html.split('class="sales-summary-table"')[1].split("</table>")[0]
        for header in ("期间", "销售大类", "终端销售收入", "销售成本",
                       "销售毛利", "毛利率"):
            self.assertIn(header, block, header)
        # 2026-07 下有纱线 / 面料两个小类 + 一行小计；2026-06 另一组
        # 小计行的「销售大类」列只写「小计（N 类）」，不再带期间
        self.assertIn("小计（2 类）", block)
        self.assertIn("小计（1 类）", block)
        self.assertNotIn("2026-07 小计", block)
        self.assertNotIn("2026-06 小计", block)
        # 7 月小计：收入 1500、成本 1200、毛利 300；合计：1700 / 1300 / 400
        self.assertIn("1,500.00", block)
        self.assertIn("1,700.00", block)
        self.assertIn("1,300.00", block)
        self.assertIn("400.00", block)

    def test_gross_profit_is_revenue_minus_cost(self):
        # 毛利口径：收入 − 成本（与损益表一致）；对不上成本的行按成本 0 计，不能漏算
        # 行1：收入 1000、成本 800 -> 毛利 200
        self._line(1, cost_price="80", cost="800", profit="200", margin="0.2")
        # 行2：收入 1000、无成本 -> 毛利 1000（以前该行不计毛利，少算）
        self._line(2)
        html = self.client.get(self.url).content.decode("utf-8")
        block = html.split('class="sales-summary-table"')[1].split("</table>")[0]
        # 合计：收入 2000、成本 800、毛利 1200
        self.assertIn("2,000.00", block)
        self.assertIn("1,200.00", block)
        self.assertIn("60.00%", block)   # 1200 / 2000

    def test_summary_period_filter_is_invoice_month(self):
        # 删掉「按出库期间 / 按开票期间」按钮，汇总表改成用「期间」筛选，
        # 而这个「期间」就是开票期间。
        line = self._line(1, out_month=self.june, cost_price="80", cost="800",
                          profit="200", margin="0.2")
        line.invoice_month = self.period.month
        line.save(update_fields=["invoice_month"])
        html = self.client.get(self.url).content.decode("utf-8")
        block = html.split('class="sales-summary-table"')[1].split("</table>")[0]
        # 出库在 6 月、开票在 7 月 -> 汇总归到 2026-07（小计行显示「小计（1 类）」）
        self.assertIn("小计（1 类）", block)
        self.assertIn("2026-07", block)
        # 汇总表标题右侧有「期间」下拉（name=invoice_month）
        seg = html[html.index("销售毛利汇总"):html.index('class="sales-summary-table"')]
        self.assertIn('name="invoice_month"', seg)
        self.assertNotIn("basis=out_month", html)
        self.assertNotIn("basis=invoice_month", html)
        # 筛选 6 月后，汇总与明细都只剩符合条件的行
        html = self.client.get(self.url, {"invoice_month": "2026-06"}).content.decode("utf-8")
        self.assertIn("没有符合筛选条件的销售台账行", html)

    def test_cost_falls_back_to_historical_price_when_no_valuation(self):
        """本期没有库存计价的物料：销售成本用上传的「历史单价」兜底（出库单价优先）。

        用户 2026-09-28：重新上传了单价，销售成本却没有变 —— 该物料本期没有月末加权价，
        成本一直是空的；现在改成取历史单价（先出库价、再入库价）。
        """
        from settlement.services.sales_ledger import fill_sales_ledger_costs

        line = self._line(1, quantity="10", revenue="1000")
        HistoricalPrice.objects.create(
            period=self.period,
            material_code="21.03",
            material_name="21.03 芳纶浆粕",
            receipt_price=Decimal("50"),
            issue_price=Decimal("80"),
        )
        updated = fill_sales_ledger_costs(self.period)
        self.assertEqual(updated, 1)
        line.refresh_from_db()
        self.assertEqual(line.cost_price, Decimal("80.00000000"))
        self.assertEqual(line.cost_amount, Decimal("800.00"))
        self.assertEqual(line.gross_profit, Decimal("200.00"))

    def test_historical_price_falls_back_to_receipt_price(self):
        # 历史单价只有入库价时，用入库价兜底
        from settlement.services.sales_ledger import fill_sales_ledger_costs

        line = self._line(1, quantity="10", revenue="1000")
        HistoricalPrice.objects.create(
            period=self.period,
            material_code="21.03",
            material_name="21.03 芳纶浆粕",
            issue_price=None,
            receipt_price=Decimal("50"),
        )
        fill_sales_ledger_costs(self.period)
        line.refresh_from_db()
        self.assertEqual(line.cost_amount, Decimal("500.00"))
        self.assertEqual(line.gross_profit, Decimal("500.00"))

    def test_month_end_valuation_price_wins_over_historical_price(self):
        # 有月末加权价时以月末加权价为准；只有取不到价时才用历史单价兜底
        from settlement.services.sales_ledger import fill_sales_ledger_costs

        line = self._line(1, quantity="10", revenue="1000")
        InventoryValuation.objects.create(
            period=self.period,
            material_code="21.03",
            material_name="21.03 芳纶浆粕",
            weighted_average_price=Decimal("90"),
        )
        HistoricalPrice.objects.create(
            period=self.period,
            material_code="21.03",
            material_name="21.03 芳纶浆粕",
            issue_price=Decimal("80"),
        )
        fill_sales_ledger_costs(self.period)
        line.refresh_from_db()
        self.assertEqual(line.cost_amount, Decimal("900.00"))

    def test_summary_title_mentions_gross_profit(self):
        # 标题只留「销售毛利汇总」（不再带「（按开票期间 · 销售大类）」）
        self._line(1, cost_price="80", cost="800", profit="200", margin="0.2")
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertIn("<h2>销售毛利汇总</h2>", html)
        self.assertNotIn("销售毛利汇总（", html)
        self.assertIn("毛利率", html)

    def test_default_invoice_month_is_current_period(self):
        # 不带参数打开 -> 期间默认当月：6 月开票的行不在默认视图里，
        # 但显式传 invoice_month= （全部）时又能看到。
        self._line(1, cost_price="80", cost="800", profit="200", margin="0.2")
        june_line = self._line(2, cost_price="80", cost="800", profit="200", margin="0.2")
        june_line.invoice_month = self.june
        june_line.save(update_fields=["invoice_month"])
        resp = self.client.get(self.url)
        self.assertEqual(resp.context["page"].paginator.count, 1)      # 默认只看当月
        self.assertEqual(resp.context["selected_invoice_month"], "2026-07")
        resp = self.client.get(self.url, {"invoice_month": ""})
        self.assertEqual(resp.context["page"].paginator.count, 2)      # 全部
        self.assertEqual(resp.context["selected_invoice_month"], "")

    def test_filter_bar_has_no_buttons_and_print_slot_next_to_channel(self):
        # 用户要求：删掉筛选栏的「筛选 / 清空」按钮，
        # 并把「打印本表」放在「销售渠道」旁边（脚本注入到 .print-here）。
        self._line(1, cost_price="80", cost="800", profit="200", margin="0.2")
        html = self.client.get(self.url).content.decode("utf-8")
        card = html[html.index('class="filter-card"'):]
        card = card[:card.index("</form>")]
        self.assertNotIn("筛选", card.split("<label>")[0])   # 按钮行已删除
        self.assertNotIn("清空", card)
        self.assertIn('class="print-here"', card)          # 但留有打印按钮容器
        # 容器紧跟在「销售渠道」之后（同一行、与下拉对齐）
        self.assertLess(card.index("销售渠道"), card.index('class="print-here"'))
        # 标题区不再放按钮，行数小字也已删
        seg = html[html.index("<h2>台账明细</h2>"):html.index('class="filter-card"')]
        self.assertNotIn("print-here", seg)
        self.assertNotIn("当前筛选", seg)

    def test_base_script_places_print_button_preferring_print_here(self):
        base = (Path(settings.BASE_DIR) / "templates" / "base.html").read_text(encoding="utf-8")
        self.assertIn("sec.querySelector('.print-here')", base)
        self.assertIn("sec.querySelector('.filter-card .filter-actions')", base)
        self.assertIn("button secondary section-print no-print", base)
        self.assertNotIn("'button small secondary section-print", base)
    def test_filter_by_invoice_status(self):
        self._line(1, cost_price="80", cost="800", profit="200", margin="0.2")
        self._line(2, status="uninvoiced", cost_price="50", cost="500")
        # 未开票行没有开票期间，默认当月会过滤掉，这里取「全部」
        resp = self.client.get(self.url, {"invoice_status": "uninvoiced",
                                           "invoice_month": ""})
        self.assertEqual(resp.status_code, 200)
        # 行数标签已从标题移除，直接用后端计数断言
        self.assertEqual(resp.context["page"].paginator.count, 1)
        self.assertIn("未开票", resp.content.decode("utf-8"))

    def test_filter_by_out_month_category_and_channel(self):
        self._line(1, cost_price="80", cost="800", profit="200", margin="0.2")
        self._line(2, out_month=self.june, category="面料", channel="国外销售",
                   cost_price="80", cost="800", profit="200", margin="0.2")
        for params in ({"out_month": "2026-06"}, {"category": "面料"},
                       {"channel": "国外销售"}):
            resp = self.client.get(self.url, params)
            self.assertEqual(resp.context["page"].paginator.count, 1, params)
        html = self.client.get(self.url, {"channel": "不存在"}).content.decode("utf-8")
        self.assertIn("没有符合筛选条件的销售台账行", html)

    def test_filter_by_invoice_month(self):
        # 开票期间：6 月开票的行只能被 2026-06 筛出
        self._line(1, cost_price="80", cost="800", profit="200", margin="0.2")
        june_line = self._line(2, cost_price="80", cost="800", profit="200", margin="0.2")
        june_line.invoice_month = self.june
        june_line.save(update_fields=["invoice_month"])
        resp = self.client.get(self.url, {"invoice_month": "2026-06"})
        self.assertEqual(resp.context["page"].paginator.count, 1)
        resp = self.client.get(self.url, {"invoice_month": "2026-07"})
        self.assertEqual(resp.context["page"].paginator.count, 1)

    def test_only_five_filter_fields_and_no_removed_ones(self):
        html = self.client.get(self.url).content.decode("utf-8")
        card = html.split('class="filter-card"')[1].split("</form>")[0]
        for kept in ("invoice_month", "out_month", "invoice_status", "category", "channel"):
            self.assertIn('name="%s"' % kept, card, kept)
        for removed in ("customer", "material", "carry_type"):
            self.assertNotIn('name="%s"' % removed, card, removed)
        self.assertEqual(card.count("<select"), 5)
        self.assertIn('class="print-here"', card)      # 打印本表容器也在筛选栏里

    def test_filter_bar_is_inside_detail_section_above_table(self):
        # 用户要求：筛选字段放在「台账明细」模块里、表格上方；不再单独占一张卡片
        html = self.client.get(self.url).content.decode("utf-8")
        # 限定在「台账明细」模块内比较（页面上还有汇总表的表格）
        detail = html[html.index("<h2>台账明细</h2>"):]
        self.assertLess(detail.index('class="filter-card"'), detail.index('class="table-wrap"'))
        self.assertLess(detail.index('name="invoice_month"'), detail.index("<table"))
        # 不再有独立的筛选卡片：全页只有一个，且在明细卡片内
        self.assertEqual(html.count('class="filter-card"'), 1)

    def test_zero_revenue_row_shows_dash_for_profit(self):
        # 已发货未开票：收入 0、只有成本，毛利为空 -> 页面显示「—」，不能出 NaN/None
        self._line(1, revenue="0", cost_price="50", cost="500")
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertNotIn("None", html.split("<tbody>")[1].split("</tbody>")[0])
        self.assertNotIn("nan", html)

    def test_export_uses_same_filter_as_page(self):
        self._line(1, cost_price="80", cost="800", profit="200", margin="0.2")
        self._line(2, status="uninvoiced", cost_price="50", cost="500")
        resp = self.client.get(
            reverse("export_listing", args=[self.period.id, "sales_ledger"]),
            {"invoice_status": "uninvoiced", "invoice_month": ""},
        )
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(BytesIO(resp.content), data_only=True)
        rows = list(wb[wb.sheetnames[0]].iter_rows(min_row=2, values_only=True))
        self.assertEqual(len(rows), 1)          # 只导出筛选后的 1 行
        self.assertEqual(rows[0][2], "未开票")

    def test_nav_has_sales_ledger_entry(self):
        base = (Path(settings.BASE_DIR) / "templates" / "base.html").read_text(
            encoding="utf-8"
        )
        self.assertIn("{% url 'sales_ledger' period.id %}", base)

class InventoryCategoryTests(SettlementFixtureMixin, TestCase):
    """库存金额结转：明细表加「物料分类」列 + 「期间 × 物料分类」数量/单位金额汇总表。"""

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="inv-cat", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.url = reverse("inventory_results", args=[self.period.id])
        self.batch = self.create_batch(
            self.period, ImportBatch.Kind.OPENING_INVENTORY, "invcat"
        )

    def _valuation(self, code, *, name="物料", oq="10", oa="100", rq="0", ra="0",
                   iq="0", ia="0", eq=None, ea=None, price="10"):
        return InventoryValuation.objects.create(
            period=self.period,
            material_code=code,
            material_name=name,
            opening_quantity=Decimal(oq),
            opening_amount=Decimal(oa),
            receipt_quantity=Decimal(rq),
            receipt_amount=Decimal(ra),
            issue_quantity=Decimal(iq),
            issue_amount=Decimal(ia),
            weighted_average_price=Decimal(price),
            ending_quantity=Decimal(eq if eq is not None else oq),
            ending_amount=Decimal(ea if ea is not None else oa),
            converged=True,
        )

    def _opening_category(self, code, category):
        return OpeningInventoryBalance.objects.create(
            period=self.period,
            batch=self.batch,
            material_code=code,
            material_name="物料",
            category_name=category,
            quantity=Decimal("1"),
            amount=Decimal("1"),
            unit_price=Decimal("1"),
        )

    def test_detail_table_has_material_category_column(self):
        self._valuation("A1")
        self._opening_category("A1", "原料")
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertIn("物料分类", html)
        self.assertIn("原料", html)

    def test_summary_table_columns_and_totals(self):
        # 原料：期初 10@100，入库 2@30；纱线：期初 5@100
        self._valuation("A1", oq="10", oa="100", rq="2", ra="30", iq="1", ia="10",
                        eq="11", ea="120")
        self._valuation("A2", oq="5", oa="100", rq="0", ra="0", iq="0", ia="0",
                        eq="5", ea="100")
        self._opening_category("A1", "原料")
        self._opening_category("A2", "纱线")
        html = self.client.get(self.url).content.decode("utf-8")
        block = html.split('class="sales-summary-table"')[1].split("</table>")[0]
        for header in ("期间", "物料分类", "期初数量", "期初金额",
                       "入库数量", "入库金额", "出库数量",
                       "出库金额", "期末数量", "期末金额"):
            self.assertIn(header, block, header)
        # 列里是金额本身，不是单位金额
        self.assertNotIn("单位金额", block)
        self.assertIn("2026-07", block)
        # 原料：期初金额 100.00、入库金额 30.00；合计期初 200.00、合计期末 220.00
        self.assertIn("100.00", block)
        self.assertIn("30.00", block)
        self.assertIn("200.00", block)
        self.assertIn("220.00", block)

    def test_category_order_and_unclassified(self):
        self._valuation("Z9", name="无分类物料")
        self._valuation("A1", name="原料物料")
        self._valuation("A2", name="纱线物料")
        self._opening_category("A1", "原料")
        self._opening_category("A2", "纱线")
        html = self.client.get(self.url).content.decode("utf-8")
        block = html.split('class="sales-summary-table"')[1].split("</table>")[0]
        self.assertLess(block.index("原料"), block.index("纱线"))
        self.assertLess(block.index("纱线"), block.index("未分类"))

    def test_export_has_category_column(self):
        self._valuation("A1")
        self._opening_category("A1", "原料")
        resp = self.client.get(
            reverse("export_listing", args=[self.period.id, "inventory_valuation"])
        )
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(BytesIO(resp.content), data_only=True)
        rows = list(wb[wb.sheetnames[0]].iter_rows(values_only=True))
        # 用户 2026-09-21：明细导出最前面加了「期间」列
        self.assertEqual(
            rows[0][:4], ("期间", "物料编码", "物料名称", "物料分类")
        )
        # 用户 2026-09-21：明细导出最前面加了「期间」列 → 物料分类落到第 4 列
        self.assertEqual(rows[1][3], "原料")

class WorkshopSettlementPageTests(SettlementFixtureMixin, TestCase):
    """车间独立结转页：不再放「导出Excel」与「打印本表」；
    「保存结转顺序」始终显示（不可用时置灰并说明原因）。"""

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="ws-page", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.url = reverse("workshop_settlements", args=[self.period.id])

    def test_no_export_excel_on_workshop_page(self):
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertNotIn("export-list/workshop_settlement", html)
        self.assertNotIn("导出Excel", html)

    def test_print_button_slot_is_switched_off(self):
        html = self.client.get(self.url).content.decode("utf-8")
        # 结转顺序板块标了 data-no-print，脚本会跳过它（不注入「打印本表」）
        self.assertIn("data-no-print", html)
        # 不再自己放「打印本表」按钮
        self.assertNotIn("打印本表</button>", html)

    def test_base_script_skips_sections_marked_no_print(self):
        base = (Path(settings.BASE_DIR) / "templates" / "base.html").read_text(encoding="utf-8")
        self.assertIn("sec.hasAttribute('data-no-print')", base)

    def test_save_order_button_is_always_visible(self):
        # 本轮以前：没有车间或已有结转时，「保存结转顺序」会被整段隐藏（用户以为功能没了）。
        # 现在：无论哪种分支都保留这个按钮，并给出当时的说明。
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertIn("保存结转顺序", html)
        self.assertTrue(
            "顺序决定上游" in html or "尚未识别到车间" in html,
            "应给出「可保存」或「暂无车间」的说明",
        )
        # 旧的条件隐藏写法已删除
        self.assertNotIn('if not has_settled and period.status != "locked"', html)

    def test_save_order_disabled_with_reason_after_settled(self):
        # 造一个「已结转」的车间，页面应把「保存结转顺序」置灰并给出原因
        WorkshopSettlement.objects.create(
            period=self.period,
            workshop="RanZheng",
            sequence=1,
            status=WorkshopSettlement.Status.SETTLED,
        )
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertIn("保存结转顺序", html)
        self.assertIn("disabled", html)
        self.assertIn("重置车间结转", html)


class PurchaseAccrualPageFilterTests(TestCase):
    """月末采购暂估页面的筛选：只保留「入库月份 / 供应商 / 物料编码」三个可见字段。

    口径（用户 2026-09-18 要求）：页面上看得见的只有这三个字段；入库单号、合同号、
    暂估状态等参数一律**不作用于页面**，避免旧链接、旧书签带上参数后出现「看不见的筛选」。
    导出 Excel 链接只带这三个参数，保证「筛选后导出 = 所见」。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="accrual-filter-user", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 8, 1))
        self.client.force_login(self.user)
        self.url = reverse("purchase_accruals", args=[self.period.id])

    def _row(self, **kwargs):
        base = {
            "receipt_month": date(2026, 8, 1),
            "document_no": "",
            "supplier_code": "S1",
            "supplier_name": "供应商一",
            "material_code": "M1",
            "material_name": "物料一",
            "contract_nos": "",
            "status": "全部暂估",
            "receipt_quantity": Decimal("1"),
            "invoiced_quantity": Decimal("0"),
            "outstanding_quantity": Decimal("1"),
            "provisional_amount": Decimal("1"),
        }
        base.update(kwargs)
        return SimpleNamespace(**base)

    def _request(self, **params):
        return SimpleNamespace(GET=params)

    def test_page_keeps_only_the_three_visible_filter_fields(self):
        resp = self.client.get(
            self.url, {"supplier": "供应商一", "status": "缺少暂估价格"}
        )
        self.assertEqual(resp.status_code, 200)
        self.assertContains(resp, 'name="supplier"')
        self.assertContains(resp, 'name="material"')
        self.assertContains(resp, 'name="receipt_month"')
        # 标题与徽标已按要求删除
        self.assertNotContains(resp, "截至月末仍未到票的入库明细")
        self.assertNotContains(resp, "供暂估应付核对")
        # 导出链接只带页面上看得见的条件，不把隐藏参数捎给导出
        export_query = resp.context["export_query"]
        self.assertIn("supplier=", export_query)
        self.assertNotIn("status=", export_query)

    def test_export_button_moved_from_heading_into_the_filter_bar(self):
        """「导出Excel」从页头挪进筛选栏（与筛选/清空/打印本表同一排），并要求按钮行样式对齐。

        用户 2026-09-18 要求：页头的「导出Excel」删掉，在筛选栏再加一个；这一排按钮的高度
        要和「入库月份/供应商/物料编码」三个框一致 —— 做法与「采购发票匹配」相同（.purchase-filter）。
        """
        resp = self.client.get(self.url, {"supplier": "供应商一"})
        html = resp.content.decode("utf-8")
        heading = re.search(r'<div class="heading-actions">(.*?)</div>', html, re.S)
        self.assertIsNotNone(heading)
        self.assertNotIn("导出Excel", heading.group(1))
        # 筛选栏里必须有导出按钮，且带上当前筛选参数
        start = html.index("filter-card purchase-filter")
        end = html.index("table-wrap", start)
        bar = html[start:end]
        self.assertIn("导出Excel", bar)
        self.assertIn("export-list/purchase_accrual/", bar)
        self.assertIn("supplier=", bar)
        # 与采购发票匹配共用同一套「筛选栏按钮 38px 对齐」样式
        self.assertIn("filter-card purchase-filter", html)

    def test_hidden_parameters_do_not_filter_the_page(self):
        rows = [
            self._row(supplier_name="甲", status="全部暂估"),
            self._row(supplier_name="乙", material_code="M2", status="缺少暂估价格"),
        ]
        self.assertEqual(
            len(views._accrual_page_filters(self._request(status="缺少暂估价格"), rows)),
            2,
        )
        self.assertEqual(
            len(views._accrual_page_filters(self._request(document_no="GR-1"), rows)), 2
        )

    def test_visible_filters_narrow_the_rows(self):
        rows = [
            self._row(receipt_month=date(2026, 6, 1), supplier_code="A",
                      supplier_name="甲供应商", material_code="M1", material_name="物料甲"),
            self._row(receipt_month=date(2026, 8, 1), supplier_code="B",
                      supplier_name="乙供应商", material_code="M2", material_name="物料乙"),
        ]
        by_month = views._accrual_page_filters(self._request(receipt_month="2026-06"), rows)
        self.assertEqual([r.material_code for r in by_month], ["M1"])
        by_supplier = views._accrual_page_filters(self._request(supplier="乙供应商"), rows)
        self.assertEqual([r.material_code for r in by_supplier], ["M2"])
        by_material = views._accrual_page_filters(self._request(material="物料甲"), rows)
        self.assertEqual([r.material_code for r in by_material], ["M1"])
        self.assertEqual(
            views._accrual_page_filters(self._request(supplier="丙供应商"), rows), []
        )

    def test_filter_labels_describe_active_filters(self):
        labels = views._accrual_filter_labels(
            self._request(receipt_month="2026-06", supplier="甲", material="M1")
        )
        self.assertEqual(len(labels), 3)
        self.assertIn("2026-06", labels[0])
        self.assertEqual(views._accrual_filter_labels(self._request()), [])

    def test_cutoff_note_states_the_as_of_date(self):
        """页面必须写明「截至哪一天」的暂估（用户 2026-09-18 要求）。

        口径依据 `build_purchase_accrual_rows`：只冲减开票期间 ≤ 本期间的发票，
        所以本页是「本期间月末」这一天的余额快照。
        """
        resp = self.client.get(self.url)
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        cutoff = views.period_month_end(self.period.month)
        self.assertEqual(cutoff, date(2026, 8, 31))
        self.assertIn(cutoff.strftime("%Y-%m-%d"), html)
        self.assertIn("暂估截止", html)
        self.assertIn("accrual-cutoff", html)
        # 说明块不能被标成不打印（.no-print 会整块隐藏）
        block = html[html.index("accrual-cutoff"):html.index("filter-card purchase-filter")]
        self.assertNotIn("no-print", block)

    def test_period_month_end_handles_short_months(self):
        self.assertEqual(views.period_month_end(date(2026, 2, 1)), date(2026, 2, 28))
        self.assertEqual(views.period_month_end(date(2028, 2, 1)), date(2028, 2, 29))
        self.assertEqual(views.period_month_end(date(2026, 4, 1)), date(2026, 4, 30))
        self.assertEqual(views.period_month_end(date(2026, 12, 1)), date(2026, 12, 31))


class PurchaseDifferencePeriodSummaryTests(SettlementFixtureMixin, TestCase):
    """暂估差异（2026-09-21 由「本月采购暂估差异」页**并入「月末采购暂估」**）。

    并入后：
    · 三张差异表（暂估差异汇总 / 按供应商物料汇总 / 价差明细）都在月末暂估页上；
    · **只显示有暂估价差的核销行** —— 暂估金额 = 发票金额（价差 0）的行、以及价差为 0 的
      期间与分组都不显示；
    · 旧地址 `/periods/<id>/purchase-differences/` 只做跳转。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="diff-period-user", password="test-password"
        )
        self.client.force_login(self.user)
        self.june = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        self.july = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.inventory_batch = self.create_batch(
            self.june, ImportBatch.Kind.INVENTORY, "inventory-diff-period"
        )
        self.contract_batch = self.create_batch(
            self.june, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-diff-period"
        )
        self.june_invoice_batch = self.create_batch(
            self.june, ImportBatch.Kind.PURCHASE_INVOICE, "invoice-diff-june"
        )
        self.july_invoice_batch = self.create_batch(
            self.july, ImportBatch.Kind.PURCHASE_INVOICE, "invoice-diff-july"
        )
        self.material = Material.objects.create(
            code="RM-DIFF-PERIOD", name="期间差异原料", uom="kg"
        )
        PurchaseContractLine.objects.create(
            period=self.june, batch=self.contract_batch, contract_no="PC-DIFF",
            contract_date=date(2026, 6, 1), supplier_code="SUP-DIFF",
            supplier_name="差异供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("100"),
            remaining_quantity=Decimal("100"), unit_price_excl_tax=Decimal("10"),
        )
        # 6 月：入库 100@10；6 月发票 60@12 → 本月入库、本月到票，价差 120
        PurchaseInvoiceLine.objects.create(
            period=self.june, batch=self.june_invoice_batch, invoice_no="PI-DIFF-06",
            invoice_date=date(2026, 6, 25), invoice_period=date(2026, 6, 1),
            contract_no="PC-DIFF", supplier_code="SUP-DIFF",
            supplier_name="差异供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("60"),
            remaining_quantity=Decimal("60"), unit_price_excl_tax=Decimal("12"),
            amount_excl_tax=Decimal("720"),
        )
        self.receipt = self.create_txn(
            self.june, self.inventory_batch, txn_date=date(2026, 6, 5),
            document_no="GR-DIFF", business_type="采购入库",
            direction=InventoryTransaction.Direction.IN, material=self.material,
            quantity=Decimal("100"), supplier_code="SUP-DIFF",
            supplier_name="差异供应商", contract_no="PC-DIFF",
        )
        auto_match_all(self.june)
        self.june.status = AccountingPeriod.Status.LOCKED
        self.june.save(update_fields=["status"])
        # 7 月：发票 40@11 核销 6 月入库 → 跨期核销，价差 40
        PurchaseInvoiceLine.objects.create(
            period=self.july, batch=self.july_invoice_batch, invoice_no="PI-DIFF-07",
            invoice_date=date(2026, 7, 20), invoice_period=date(2026, 7, 1),
            contract_no="PC-DIFF", supplier_code="SUP-DIFF",
            supplier_name="差异供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("40"),
            remaining_quantity=Decimal("40"), unit_price_excl_tax=Decimal("11"),
            amount_excl_tax=Decimal("440"),
        )
        auto_match_all(self.july)
        self.url = reverse("purchase_accruals", args=[self.july.id])

    def test_summary_table_has_period_column_and_grand_total(self):
        resp = self.client.get(self.url)
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        # 顶部卡片排已改成表格
        self.assertNotIn('class="summary-strip"', html)
        self.assertIn('class="difference-summary-table"', html)
        block = html.split('class="difference-summary-table"')[1].split("</table>")[0]
        for header in ("期间", "本月核销数量", "暂估金额", "发票金额",
                       "暂估价差", "跨期核销"):
            self.assertIn(header, block, header)
        self.assertIn("合计", block)

        rows = resp.context["diff_summary_rows"]
        self.assertEqual([str(row["period"]) for row in rows],
                         [str(self.june), str(self.july)])
        june, july = rows
        # 6 月：本月入库 + 本月到票
        self.assertEqual(june["matched_quantity"], Decimal("60"))
        self.assertEqual(june["provisional_amount"], Decimal("600.00"))
        self.assertEqual(june["invoice_amount"], Decimal("720.00"))
        self.assertEqual(june["difference_amount"], Decimal("120.00"))
        self.assertEqual(june["cross_period_count"], 0)
        self.assertFalse(june["current"])
        # 7 月：跨期核销 6 月入库
        self.assertEqual(july["matched_quantity"], Decimal("40"))
        self.assertEqual(july["provisional_amount"], Decimal("400.00"))
        self.assertEqual(july["invoice_amount"], Decimal("440.00"))
        self.assertEqual(july["difference_amount"], Decimal("40.00"))
        self.assertEqual(july["cross_period_count"], 1)
        self.assertTrue(july["current"])
        # 合计 = 各期间之和，也与当期明细小计一致
        total = resp.context["diff_summary_total"]
        self.assertEqual(total["matched_quantity"], Decimal("100"))
        self.assertEqual(total["provisional_amount"], Decimal("1000.00"))
        self.assertEqual(total["invoice_amount"], Decimal("1160.00"))
        self.assertEqual(total["difference_amount"], Decimal("160.00"))
        self.assertEqual(total["cross_period_count"], 1)
        self.assertEqual(total["pending_count"], 0)
        for key in ("matched_quantity", "provisional_amount", "invoice_amount",
                    "difference_amount", "cross_period_count", "pending_count"):
            self.assertEqual(sum(row[key] for row in rows), total[key], key)
        totals = resp.context["diff_totals"]
        self.assertEqual(july["difference_amount"], totals["difference_amount"])

    def test_difference_detail_uses_period_column_without_invoice_no(self):
        """用户 2026-09-21：暂估价差明细把「发票」列改成「期间」，且不再显示发票号。"""
        html = self.client.get(self.url).content.decode("utf-8")
        seg = html.split("<h2>暂估价差明细</h2>", 1)[1].split("</section>", 1)[0]
        heads = [re.sub(r"<[^>]+>", "", h).strip()
                 for h in re.findall(r"<th[^>]*>(.*?)</th>",
                                     seg.split("</thead>")[0], re.S)]
        self.assertEqual(heads[0], "期间")
        self.assertNotIn("发票号", heads)
        self.assertNotIn("发票", heads)
        # 明细里不再出现发票编号（B166-xx 这类）
        self.assertNotIn("PI-DIFF-07", seg)
        self.assertNotIn("source_document_detail", seg)
        # 第一格是发票所属期间
        first = re.findall(r"<td[^>]*>(.*?)</td>",
                           seg.split("</thead>", 1)[1].split("</tr>", 1)[0], re.S)
        self.assertEqual(re.sub(r"<[^>]+>", "", first[0]).strip(), "2026-07")

    def test_period_table_follows_page_filters(self):
        """筛选后各期间汇总也跟着变（与明细同源）；价差变 0 的期间直接不显示。"""
        resp = self.client.get(self.url, {"cross_period": "yes"})
        rows = resp.context["diff_summary_rows"]
        # 只看跨期核销 → 6 月没有跨期行、价差为 0 → 该期间整行不显示
        self.assertEqual([str(row["period"]) for row in rows], [str(self.july)])
        self.assertEqual(rows[0]["difference_amount"], Decimal("40.00"))
        self.assertEqual(resp.context["diff_summary_total"]["difference_amount"],
                         Decimal("40.00"))
        resp2 = self.client.get(self.url, {"supplier": "查无此供应商"})
        self.assertEqual(resp2.context["diff_summary_rows"], [])
        self.assertEqual(resp2.context["diff_summary_total"]["difference_amount"],
                         Decimal("0"))

    def test_rows_without_difference_are_hidden(self):
        """用户 2026-09-21：没有暂估价差的核销行不显示（页面与汇总都一样）。"""
        # 7 月再造一笔「票价 = 暂估价（10）」的采购入库 + 到票 → 价差 0
        material = Material.objects.create(code="RM-NODIFF", name="无价差原料", uom="kg")
        PurchaseContractLine.objects.create(
            period=self.july, batch=self.contract_batch, contract_no="PC-NODIFF",
            contract_date=date(2026, 7, 1), supplier_code="SUP-NODIFF",
            supplier_name="无价差供应商", material_code=material.code,
            material_name=material.name, quantity=Decimal("50"),
            remaining_quantity=Decimal("50"), unit_price_excl_tax=Decimal("10"),
        )
        self.create_txn(
            self.july, self.inventory_batch, txn_date=date(2026, 7, 6),
            document_no="GR-NODIFF", business_type="采购入库",
            direction=InventoryTransaction.Direction.IN, material=material,
            quantity=Decimal("50"), supplier_code="SUP-NODIFF",
            supplier_name="无价差供应商", contract_no="PC-NODIFF",
        )
        PurchaseInvoiceLine.objects.create(
            period=self.july, batch=self.july_invoice_batch, invoice_no="PI-NODIFF",
            invoice_date=date(2026, 7, 25), invoice_period=date(2026, 7, 1),
            contract_no="PC-NODIFF", supplier_code="SUP-NODIFF",
            supplier_name="无价差供应商", material_code=material.code,
            material_name=material.name, quantity=Decimal("50"),
            remaining_quantity=Decimal("50"), unit_price_excl_tax=Decimal("10"),
            amount_excl_tax=Decimal("500"),
        )
        auto_match_all(self.july)

        resp = self.client.get(self.url)
        html = resp.content.decode("utf-8")
        # 明细里没有这笔无价差的核销
        self.assertEqual(
            [row.material_code for row in resp.context["diff_page"]], ["RM-DIFF-PERIOD"]
        )
        # 暂估差异三块（暂估差异汇总 / 暂估价差汇总 / 暂估价差明细）里不该出现它；
        # 注意：范围要收在「回冲明细（逐笔）」之前 —— 那张表列的是全部发票核销，
        # 本来就该有这笔无价差核销（它不属于"价差"表）。
        diff_region = html.split("<h2>暂估差异汇总</h2>", 1)[1].split(
            "<h2>回冲明细（逐笔）</h2>", 1
        )[0]
        self.assertNotIn("PI-NODIFF", diff_region)
        self.assertNotIn("无价差供应商", diff_region)
        # 汇总也不含它
        self.assertEqual(
            [row["material_code"] for row in resp.context["diff_grouped_rows"]],
            ["RM-DIFF-PERIOD"],
        )
        # 价差合计不变（仍是 7 月那 40）
        self.assertEqual(resp.context["diff_totals"]["difference_amount"],
                         Decimal("40.00"))


class PurchaseFlowSummaryTests(SettlementFixtureMixin, TestCase):
    """「当月发生额」四件套表：本月暂估 / 本月到票 / 当月到货到票 / 本月回冲金额。

    位置（用户 2026-09-18 指定）：**月末采购暂估页**（先前放「回冲汇总」，已按要求搬过来）。
    「本月到票」＝本月发票核销金额合计，且恒等于 当月到货到票 ＋ 本月回冲金额
    （回冲＝核销以前月份入库/期初台账的部分）。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="flow-summary-user", password="test-password"
        )
        self.client.force_login(self.user)
        self.june = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        self.july = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.inventory_batch = self.create_batch(
            self.june, ImportBatch.Kind.INVENTORY, "inventory-flow"
        )
        self.contract_batch = self.create_batch(
            self.june, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-flow"
        )
        self.june_invoice_batch = self.create_batch(
            self.june, ImportBatch.Kind.PURCHASE_INVOICE, "invoice-flow-june"
        )
        self.july_invoice_batch = self.create_batch(
            self.july, ImportBatch.Kind.PURCHASE_INVOICE, "invoice-flow-july"
        )
        self.material = Material.objects.create(
            code="RM-FLOW", name="四件套原料", uom="kg"
        )
        PurchaseContractLine.objects.create(
            period=self.june, batch=self.contract_batch, contract_no="PC-FLOW",
            contract_date=date(2026, 6, 1), supplier_code="SUP-FLOW",
            supplier_name="四件套供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("100"),
            remaining_quantity=Decimal("100"), unit_price_excl_tax=Decimal("10"),
        )
        # 6 月：入库 100@10 ＋ 6 月发票 60@12（当月到货到票），剩 40 未到票（本月暂估 400）
        PurchaseInvoiceLine.objects.create(
            period=self.june, batch=self.june_invoice_batch, invoice_no="PI-FLOW-06",
            invoice_date=date(2026, 6, 25), invoice_period=date(2026, 6, 1),
            contract_no="PC-FLOW", supplier_code="SUP-FLOW",
            supplier_name="四件套供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("60"),
            remaining_quantity=Decimal("60"), unit_price_excl_tax=Decimal("12"),
            amount_excl_tax=Decimal("720"),
        )
        self.receipt = self.create_txn(
            self.june, self.inventory_batch, txn_date=date(2026, 6, 5),
            document_no="GR-FLOW", business_type="采购入库",
            direction=InventoryTransaction.Direction.IN, material=self.material,
            quantity=Decimal("100"), supplier_code="SUP-FLOW",
            supplier_name="四件套供应商", contract_no="PC-FLOW",
        )
        auto_match_all(self.june)
        self.june.status = AccountingPeriod.Status.LOCKED
        self.june.save(update_fields=["status"])
        # 7 月：发票 40@11 核销 6 月入库 → 本月回冲 440
        PurchaseInvoiceLine.objects.create(
            period=self.july, batch=self.july_invoice_batch, invoice_no="PI-FLOW-07",
            invoice_date=date(2026, 7, 20), invoice_period=date(2026, 7, 1),
            contract_no="PC-FLOW", supplier_code="SUP-FLOW",
            supplier_name="四件套供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("40"),
            remaining_quantity=Decimal("40"), unit_price_excl_tax=Decimal("11"),
            amount_excl_tax=Decimal("440"),
        )
        auto_match_all(self.july)
        self.url = reverse("purchase_accruals", args=[self.july.id])
        # 「暂估回冲」已并入本页（2026-09-21）：旧 URL 只做跳转，保留用于跳转断言
        self.old_summary_url = reverse("purchase_summary", args=[self.july.id])

    def test_flow_table_columns_and_values(self):
        resp = self.client.get(self.url)
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        self.assertIn('class="purchase-flow-table"', html)
        block = html.split('class="purchase-flow-table"')[1].split("</table>")[0]
        # 列顺序：期初暂估 → 合同金额（用户 2026-09-19 指定，加在期初暂估与本月到票中间）
        # → 本月到票 → 当月到货到票 → 本月暂估 → 本月回冲 → 期末暂估
        self.assertEqual(
            re.findall(r"<th[^>]*>([^<]+)</th>", block.split("</thead>")[0]),
            # 用户 2026-09-21：本月回冲之后、期末暂估之前新增「暂估价差」
            ["期间", "期初暂估", "合同金额", "本月到票", "当月到货到票",
             "本月暂估", "本月回冲", "暂估价差", "期末暂估"],
        )
        self.assertIn("合计", block)
        # 标题为「到票与暂估」，且「各期间月末采购暂估」模块已按要求删除
        self.assertIn("<h2>到票与暂估</h2>", html)
        self.assertNotIn("各期间月末采购暂估", html)
        self.assertNotIn("accrual-summary-table", html)
        self.assertNotIn("当月发生额", block)
        self.assertNotIn(" 笔", block)
        self.assertNotIn(" 行", block)
        # 口径说明（用户 2026-09-19 重写、2026-09-21 全站统一）：
        # 左上角「口径说明」标签 + 一个字段一行，列＝字段 / 数据来源 / 计算依据 / 作用
        self.assertIn('class="basis-title">口径说明<', html)
        self.assertIn('class="basis-table"', html)
        self.assertNotIn('class="accrual-basis"', html)   # 旧的长段落 ul 已删除
        for label in ("数据来源", "计算依据", "作用"):
            self.assertIn("<th>%s</th>" % label, html, label)
        basis = html.split('class="basis-table"', 1)[1].split("</table>", 1)[0]
        for column in ("期初暂估", "合同金额", "本月到票", "当月到货到票",
                       "本月暂估", "本月回冲", "期末暂估"):
            self.assertIn("<strong>%s</strong>" % column, basis, column)
        # 每个字段都要同时给出「来源 / 依据 / 作用」三格（8 行 × 4 格；含 2026-09-21 新增的「暂估价差」）
        self.assertEqual(basis.count("<tr>"), 9)          # 1 行表头 + 8 行字段
        self.assertEqual(basis.count("<td>"), 32)
        self.assertIn("<strong>暂估价差</strong>", basis)
        # 表下保留一段简短注解：归属月份判定 + 余额不可加 + 回冲是发票金额口径
        note = html.split("</table>", 1)[1].split("</section>", 1)[0]
        self.assertIn("归属月份", note)
        self.assertIn("不可加", note)
        # 未被删掉的部分：暂估截止说明 + 未到票明细
        self.assertIn("accrual-cutoff", html)

        rows = {str(r["period"]): r for r in resp.context["flow_rows"]}
        june = rows[str(self.june)]
        july = rows[str(self.july)]
        # 6 月：期初 0；当月到货到票 720（60@12）；本月暂估 400（剩 40@10）；期末暂估 400
        self.assertEqual(june["opening_accrual"], Decimal("0.00"))
        # 合同金额（用户 2026-09-19 新增列）：归属本期的合同不含税金额 ＝ 100 × 10
        self.assertEqual(june["contract_amount"], Decimal("1000.00"))
        self.assertEqual(july["contract_amount"], Decimal("0.00"))
        self.assertEqual(
            resp.context["flow_total"]["contract_amount"], Decimal("1000.00")
        )
        self.assertEqual(june["same_month"], Decimal("720.00"))
        self.assertEqual(june["month_accrual"], Decimal("400.00"))
        self.assertEqual(june["invoice_matched"], Decimal("720.00"))
        self.assertEqual(june["cross"], Decimal("0.00"))
        self.assertEqual(june["ending_accrual"], Decimal("400.00"))
        self.assertFalse(june["current"])
        # 7 月：期初＝6 月期末 400；本月无入库；发票核销的是 6 月入库 → 全部算回冲；期末暂估 0
        # 「暂估价差」列（用户 2026-09-21）：＝ 该期发票核销产生的价差合计（发票金额 − 暂估金额）
        # 6 月：60@12 核销 60@10 → 120；7 月：40@11 核销 40@10 → 40；合计 160
        self.assertEqual(june["difference_amount"], Decimal("120.00"))
        self.assertEqual(july["difference_amount"], Decimal("40.00"))
        self.assertEqual(resp.context["flow_total"]["difference_amount"], Decimal("160.00"))
        self.assertEqual(july["opening_accrual"], Decimal("400.00"))
        self.assertEqual(july["same_month"], Decimal("0.00"))
        self.assertEqual(july["month_accrual"], Decimal("0.00"))
        self.assertEqual(july["invoice_matched"], Decimal("440.00"))
        self.assertEqual(july["cross"], Decimal("440.00"))
        self.assertEqual(july["ending_accrual"], Decimal("0.00"))
        # 期初（本期）＝ 上一期期末
        self.assertEqual(july["opening_accrual"], june["ending_accrual"])
        # 期末暂估 ＝ 该期月末未到票余额（build_purchase_accrual_rows 口径）
        self.assertEqual(
            june["ending_accrual"],
            sum((r.provisional_amount or Decimal("0")
                 for r in build_purchase_accrual_rows(self.june)), Decimal("0")),
        )
        # 合计行：期初取第一期、期末取最后一期（余额不可加）
        self.assertEqual(resp.context["flow_total"]["opening_accrual"],
                         june["opening_accrual"])
        self.assertEqual(resp.context["flow_total"]["ending_accrual"],
                         july["ending_accrual"])

    def test_merged_page_has_flow_table_and_reversal_detail(self):
        """「暂估回冲」并入「月末暂估」（2026-09-21）：一页里既有「到票与暂估」也有「回冲明细（逐笔）」，
        旧 URL 只做跳转。"""
        resp = self.client.get(self.url)
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        self.assertIn("purchase-flow-table", html)
        self.assertIn("<h2>到票与暂估</h2>", html)
        self.assertIn("<h2>回冲明细（逐笔）</h2>", html)
        # 用户 2026-09-21：整页的「明细对照」已删除
        self.assertNotIn("明细对照", html)
        # 旧 URL → 跳到月末暂估（查询参数原样带过去）
        old = self.client.get(self.old_summary_url)
        self.assertEqual(old.status_code, 302)
        self.assertEqual(old["Location"], self.url)
        old2 = self.client.get(self.old_summary_url, {"supplier": "四件套"})
        self.assertEqual(old2.status_code, 302)
        self.assertEqual(old2["Location"], self.url + "?supplier=%E5%9B%9B%E4%BB%B6%E5%A5%97")

    def test_reversal_detail_merged_into_summary_page(self):
        """「回冲明细」已并入「暂估回冲」：一页里既有汇总也有本期逐笔明细（2026-09-19）。"""
        resp = self.client.get(self.url)
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        self.assertIn("<h2>回冲明细（逐笔）</h2>", html)
        self.assertNotIn("<h2>明细对照</h2>", html)
        # 逐笔明细＝本期核销：7 月只有 1 笔（40@11 核销 6 月入库，属跨期）
        rows = resp.context["detail_rows"]
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["invoice_no"], "PI-FLOW-07")
        self.assertTrue(rows[0]["cross_period"])
        totals = resp.context["detail_totals"]
        self.assertEqual(totals["count"], 1)
        self.assertEqual(totals["matched_quantity"], Decimal("40"))
        self.assertEqual(totals["provisional_amount"], Decimal("400.00"))
        self.assertEqual(totals["invoice_amount"], Decimal("440.00"))
        self.assertEqual(totals["difference_amount"], Decimal("40.00"))
        self.assertEqual(totals["cross_period_count"], 1)
        # 逐期间汇总表：6 月 1 笔＋7 月 1 笔，各算各的
        periods = resp.context["reversal_period_rows"]
        self.assertEqual([str(r["period"]) for r in periods],
                         [str(self.june), str(self.july)])
        self.assertEqual(resp.context["reversal_period_total"]["count"], 2)
        self.assertEqual(
            resp.context["reversal_period_total"]["invoice_amount"], Decimal("1160.00")
        )
        # 导出 + 打印落点都在「回冲明细（逐笔）」板块里；
        # 该板块自己的筛选栏已按用户要求删除（页面主筛选栏是月末暂估明细的，与本板块无关）
        seg = html.split("<h2>回冲明细（逐笔）</h2>", 1)[1].split("</section>", 1)[0]
        for field in ("invoice_no", "supplier", "material", "cross_period"):
            self.assertNotIn('name="%s"' % field, seg, field)
        self.assertIn(reverse("export_listing", args=[self.july.id, "purchase_reversal"]), seg)
        self.assertIn('class="print-here"', seg)

    def test_reversal_summary_is_a_period_table(self):
        """回冲明细汇总＝带「期间」列的表格（一个期间一行，多期间才有合计行），不是卡片排。"""
        resp = self.client.get(self.url)
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        self.assertIn("<h1>月末采购暂估</h1>", html)
        self.assertNotIn("个期间 · 累计口径", html)
        self.assertNotIn("累计口径", html)
        self.assertNotIn('class="summary-strip"', html)
        block = html.split('class="reversal-summary-table"')[1].split("</table>")[0]
        self.assertEqual(
            re.findall(r"<th[^>]*>([^<]+)</th>", block.split("</thead>")[0]),
            ["期间", "回冲笔数", "核销数量", "暂估金额", "发票金额", "价差合计", "跨期回冲"],
        )
        self.assertIn("合计", block)
        rows = resp.context["reversal_period_rows"]
        self.assertEqual([str(r["period"]) for r in rows],
                         [str(self.june), str(self.july)])
        self.assertFalse(rows[0]["current"])
        self.assertTrue(rows[1]["current"])
        # 6 月：核销的是 6 月自己的入库（本期，不算跨期）
        self.assertEqual(rows[0]["count"], 1)
        self.assertEqual(rows[0]["invoice_amount"], Decimal("720.00"))
        self.assertEqual(rows[0]["cross_period_count"], 0)
        # 7 月：1 笔跨期回冲
        self.assertEqual(rows[1]["count"], 1)
        self.assertEqual(rows[1]["matched_quantity"], Decimal("40"))
        self.assertEqual(rows[1]["invoice_amount"], Decimal("440.00"))
        self.assertEqual(rows[1]["cross_period_count"], 1)
        total = resp.context["reversal_period_total"]
        self.assertEqual(total["count"], 2)
        self.assertEqual(total["invoice_amount"], Decimal("1160.00"))
        self.assertEqual(total["cross_period_count"], 1)

    def test_detail_match_module_removed(self):
        """2026-09-21：用户认为「明细对照」意义不大、且与「各期间对比」重复 —— 整块删除。
        保留「各期间对比」+ 它的口径说明表（字段 / 数据来源 / 计算依据 / 作用）。"""
        resp = self.client.get(self.url)
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        self.assertNotIn("明细对照", html)
        self.assertNotIn("detail-match-table", html)
        self.assertNotIn("detail-grid", html)
        self.assertNotIn("detail-col", html)
        # 已删除的整页断言（合并后不再是只有两块）
        # 明细对照里的那几行也一并消失（这些词在别处不会出现：
        # 「期初暂估台账 / 暂估合计」等仍作为各期间对比的表头或口径字段保留，故不在此列）
        for label in ("上传合计", "已核销（配上货）", "未核销", "回冲金额", "其中跨期核销"):
            self.assertNotIn(label, html, label)
        # 口径说明：页面上有多张（到票与暂估 / 暂估差异），都带左上角标签
        self.assertIn('class="basis-title">口径说明<', html)

    def test_detail_table_drops_invoice_no_column(self):
        """用户 2026-09-21：「回冲明细（逐笔）」那张明细表不再显示「发票号」列。

        表头由 21 列变 20 列，每行单元格数也要跟着变（否则表格会错位）。
        导出 Excel 里仍保留发票号，便于与发票台账对账。
        """
        html = self.client.get(self.url).content.decode("utf-8")
        seg = html.split("<h2>回冲明细（逐笔）</h2>", 1)[1].split("</section>", 1)[0]
        rest = seg.split('class="reversal-summary-table"', 1)[1].split("</table>", 1)[1]
        table = rest[:rest.index("</table>")]
        heads = re.findall(r"<th[^>]*>([^<]+)</th>", table.split("</thead>")[0])
        self.assertNotIn("发票号", heads)
        self.assertEqual(len(heads), 20)
        self.assertEqual(heads[0], "发票日期")
        rows = re.findall(r"<tr>(.*?)</tr>", table.split("</thead>", 1)[1], re.S)
        self.assertTrue(rows)
        for row in rows:
            self.assertEqual(len(re.findall(r"<td[^>]*>", row)), 20, row[:80])
        # 空态那一行的 colspan 也要跟着改成 20（用一个筛不到任何记录的筛选条件触发）
        empty_html = self.client.get(
            self.url, {"supplier": "没有这个供应商"}
        ).content.decode("utf-8")
        self.assertIn('colspan="20"', empty_html)
        self.assertIn("没有符合筛选条件的回冲记录", empty_html)

    def test_old_reversal_page_redirects_to_summary(self):
        """旧的「回冲明细」URL 变成跳转：旧链接/书签照样能用，查询参数原样带走。"""
        old = reverse("purchase_reversal_report", args=[self.july.id])
        resp = self.client.get(old)
        self.assertEqual(resp.status_code, 302)
        self.assertEqual(resp["Location"], self.url)
        resp = self.client.get(old, {"supplier": "四件套"})
        self.assertEqual(resp.status_code, 302)
        self.assertTrue(resp["Location"].startswith(self.url + "?"))
        self.assertIn("supplier=", resp["Location"])

    def test_detail_filter_only_narrows_the_detail_table(self):
        """明细筛选（现只剩老链接里的参数）只影响下方逐笔表；逐期间汇总表仍是全量。"""
        full = self.client.get(self.url).context
        filtered = self.client.get(self.url, {"cross_period": "false"}).context
        self.assertEqual(filtered["detail_rows"], [])
        self.assertEqual(filtered["detail_totals"]["count"], 0)
        self.assertEqual(
            [(r["count"], r["invoice_amount"]) for r in filtered["reversal_period_rows"]],
            [(r["count"], r["invoice_amount"]) for r in full["reversal_period_rows"]],
        )

    def test_earlier_months_keep_their_numbers_when_a_later_month_is_added(self):
        """到以后月份时，以前月份那一行的数据必须保持不变（按各期当时口径取数）。"""
        # 先记下 6 / 7 月的数
        before = {
            str(row["period"]): {
                key: row[key]
                for key in ("opening_accrual", "invoice_matched", "same_month",
                            "month_accrual", "cross", "ending_accrual")
            }
            for row in self.client.get(self.url).context["flow_rows"]
        }
        # 新增 8 月：入库 50@20 + 当月发票 20@25（当月到货到票），并核销 6 月剩余的 40 里的一部分
        august = AccountingPeriod.objects.create(month=date(2026, 8, 1))
        inventory_batch = self.create_batch(
            august, ImportBatch.Kind.INVENTORY, "inventory-flow-august"
        )
        invoice_batch = self.create_batch(
            august, ImportBatch.Kind.PURCHASE_INVOICE, "invoice-flow-august"
        )
        material2 = Material.objects.create(code="RM-FLOW-AUG", name="八月原料", uom="kg")
        PurchaseContractLine.objects.create(
            period=august, batch=self.create_batch(
                august, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-flow-august"),
            contract_no="PC-FLOW-AUG", contract_date=date(2026, 8, 1),
            supplier_code="SUP-FLOW", supplier_name="四件套供应商",
            material_code=material2.code, material_name=material2.name,
            quantity=Decimal("50"), remaining_quantity=Decimal("50"),
            unit_price_excl_tax=Decimal("20"),
        )
        PurchaseInvoiceLine.objects.create(
            period=august, batch=invoice_batch, invoice_no="PI-FLOW-08",
            invoice_date=date(2026, 8, 20), invoice_period=date(2026, 8, 1),
            contract_no="PC-FLOW-AUG", supplier_code="SUP-FLOW",
            supplier_name="四件套供应商", material_code=material2.code,
            material_name=material2.name, quantity=Decimal("20"),
            remaining_quantity=Decimal("20"), unit_price_excl_tax=Decimal("25"),
            amount_excl_tax=Decimal("500"),
        )
        self.create_txn(
            august, inventory_batch, txn_date=date(2026, 8, 5),
            document_no="GR-FLOW-AUG", business_type="采购入库",
            direction=InventoryTransaction.Direction.IN, material=material2,
            quantity=Decimal("50"), supplier_code="SUP-FLOW",
            supplier_name="四件套供应商", contract_no="PC-FLOW-AUG",
        )
        auto_match_all(august)

        resp = self.client.get(reverse("purchase_accruals", args=[august.id]))
        rows = {str(row["period"]): row for row in resp.context["flow_rows"]}
        self.assertIn(str(august), rows)
        # 6 / 7 月那一行一个数字都没变
        for period_key, expected in before.items():
            for key, value in expected.items():
                self.assertEqual(rows[period_key][key], value,
                                 "%s %s" % (period_key, key))
        # 8 月自己的数（期初＝7 月期末＝0；当月到货到票 20@25=500；入库 50 只到票 20 → 本月暂估 30@20=600）
        self.assertEqual(rows[str(august)]["opening_accrual"], Decimal("0.00"))
        self.assertEqual(rows[str(august)]["same_month"], Decimal("500.00"))
        self.assertEqual(rows[str(august)]["month_accrual"], Decimal("600.00"))
        self.assertEqual(rows[str(august)]["ending_accrual"], Decimal("600.00"))

    def test_month_invoice_equals_same_month_plus_reversal(self):
        """恒等式：本月到票 ＝ 当月到货到票 ＋ 本月回冲（逐期间成立）。"""
        resp = self.client.get(self.url)
        for row in resp.context["flow_rows"]:
            self.assertEqual(
                row["invoice_matched"], row["same_month"] + row["cross"],
                str(row["period"]),
            )
        current = [r for r in resp.context["flow_rows"] if r["current"]][0]
        # 本月暂估与「月末采购暂估」口径一致：= 本月入库行的暂估金额之和
        # （上一张「各期间月末采购暂估」的金额是"截至月末余额"，含以前月份，所以两者不等）
        month_rows = [
            r for r in build_purchase_accrual_rows(self.july)
            if r.receipt_month == self.july.month
        ]
        self.assertEqual(
            current["month_accrual"],
            sum((r.provisional_amount or Decimal("0") for r in month_rows),
                Decimal("0")),
        )
        # 合计 = 各期间之和
        total = resp.context["flow_total"]
        for key in ("same_month", "month_accrual", "invoice_matched", "cross"):
            self.assertEqual(sum(r[key] for r in resp.context["flow_rows"]),
                             total[key], key)
        self.assertEqual(total["invoice_matched"], Decimal("1160.00"))
        self.assertEqual(total["same_month"], Decimal("720.00"))
        self.assertEqual(total["cross"], Decimal("440.00"))

    def test_ending_equals_opening_plus_month_accrual_minus_reversal(self):
        """余额滚动：期末暂估 ＝ 期初暂估 ＋ 本月暂估 − 该月核销掉的以前暂估金额。

        （本月回冲列是发票金额口径，与此处的暂估金额差一个价差，见页面「注意」一行。）
        """
        resp = self.client.get(self.url)
        rows = {str(r["period"]): r for r in resp.context["flow_rows"]}
        for item in (self.june, self.july):
            row = rows[str(item)]
            cross_provisional = sum(
                (d.provisional_amount or Decimal("0")
                 for d in build_purchase_difference_rows(item) if d.cross_period),
                Decimal("0"),
            )
            self.assertEqual(
                row["ending_accrual"],
                row["opening_accrual"] + row["month_accrual"] - cross_provisional,
                str(item),
            )
class PurchaseAccrualDetailTablesTests(SettlementFixtureMixin, TestCase):
    """「到票与暂估」表下方 5 个明细块：合计必须等于**本期那一行**的数字。

    背景：明细块是 2026-09-18 收尾时加上的（当时只看了渲染、没补测试）。
    2026-09-19 补测试时发现一个只在多期间下暴露的口径错：明细块是「本期」口径，
    而徽标当时取的是 flow_total（跨期间合计；期初/期末取首末期），
    于是两个期间时「本月回冲」徽标会显示两期之和、4 月页面的「期初暂估」徽标会显示首期期初。
    已改为 flow_current，本用例锁住这一点。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="accrual-detail-user", password="test-password"
        )
        self.client.force_login(self.user)
        self.june = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        self.july = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.june_invoice_batch = self.create_batch(
            self.june, ImportBatch.Kind.PURCHASE_INVOICE, "invoice-detail-june"
        )
        self.july_invoice_batch = self.create_batch(
            self.july, ImportBatch.Kind.PURCHASE_INVOICE, "invoice-detail-july"
        )
        self.contract_batch = self.create_batch(
            self.june, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-detail"
        )
        self.inventory_batch = self.create_batch(
            self.june, ImportBatch.Kind.INVENTORY, "inventory-detail"
        )
        self.material = Material.objects.create(
            code="RM-DETAIL", name="明细原料", uom="kg"
        )
        PurchaseContractLine.objects.create(
            period=self.june, batch=self.contract_batch, contract_no="PC-DETAIL",
            contract_date=date(2026, 6, 1), supplier_code="SUP-DETAIL",
            supplier_name="明细供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("100"),
            remaining_quantity=Decimal("100"), unit_price_excl_tax=Decimal("10"),
        )
        # 6 月：入库 100@10；发票 60@12 核销本月入库 → 当月到货到票 720，余 40 未到票 → 本月暂估 400
        PurchaseInvoiceLine.objects.create(
            period=self.june, batch=self.june_invoice_batch, invoice_no="PI-DETAIL-06",
            invoice_date=date(2026, 6, 25), invoice_period=date(2026, 6, 1),
            contract_no="PC-DETAIL", supplier_code="SUP-DETAIL",
            supplier_name="明细供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("60"),
            remaining_quantity=Decimal("60"), unit_price_excl_tax=Decimal("12"),
            amount_excl_tax=Decimal("720"),
        )
        self.receipt = self.create_txn(
            self.june, self.inventory_batch, txn_date=date(2026, 6, 5),
            document_no="GR-DETAIL", business_type="采购入库",
            direction=InventoryTransaction.Direction.IN, material=self.material,
            quantity=Decimal("100"), supplier_code="SUP-DETAIL",
            supplier_name="明细供应商", contract_no="PC-DETAIL",
        )
        auto_match_all(self.june)
        self.june.status = AccountingPeriod.Status.LOCKED
        self.june.save(update_fields=["status"])
        # 7 月：发票 40@11 核销 6 月入库 → 本月回冲 440
        PurchaseInvoiceLine.objects.create(
            period=self.july, batch=self.july_invoice_batch, invoice_no="PI-DETAIL-07",
            invoice_date=date(2026, 7, 20), invoice_period=date(2026, 7, 1),
            contract_no="PC-DETAIL", supplier_code="SUP-DETAIL",
            supplier_name="明细供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("40"),
            remaining_quantity=Decimal("40"), unit_price_excl_tax=Decimal("11"),
            amount_excl_tax=Decimal("440"),
        )
        auto_match_all(self.july)

    @staticmethod
    def _pick(row, key):
        if isinstance(row, dict):
            return row.get(key)
        return getattr(row, key, None)

    def _sum(self, rows, key):
        return sum((self._pick(row, key) or Decimal("0") for row in rows), Decimal("0"))

    def _pill(self, html, heading):
        # 徽标位置随表格变化（用户 2026-09-19）：4 张有筛选栏的表在筛选栏那一排，
        # 期初暂估明细仍在标题右侧。这里只在这个区块内找第一个「名称 金额」徽标。
        seg = html.split("<h2>%s</h2>" % heading, 1)[1].split("</section>", 1)[0]
        m = re.search(r'<span class="pill[^"]*">[^<]*?([\d,]+\.\d\d)', seg)
        self.assertIsNotNone(m, heading)
        return m.group(1)

    def test_five_detail_blocks_are_rendered(self):
        """5 个明细块 + 各自的口径说明都在页面上。"""
        resp = self.client.get(reverse("purchase_accruals", args=[self.july.id]))
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        for heading in ("到票与暂估", "期初暂估明细", "本月到票明细",
                        "当月到货到票明细", "本月暂估明细", "本月回冲明细"):
            self.assertIn("<h2>%s</h2>" % heading, html)
        self.assertGreaterEqual(html.count("口径："), 5)

    def test_detail_totals_equal_current_period_row(self):
        """明细合计 = 本期那一行（不是跨期间合计）。"""
        resp = self.client.get(reverse("purchase_accruals", args=[self.july.id]))
        details = resp.context["details"]
        current = [r for r in resp.context["flow_rows"] if r["current"]][0]
        total = resp.context["flow_total"]
        # 7 月那一行：期初＝6 月期末 400；到票 440 且全是回冲；本月无入库 → 本月暂估 0
        self.assertEqual(self._sum(details["opening"], "amount"),
                         current["opening_accrual"])
        self.assertEqual(self._sum(details["opening"], "amount"), Decimal("400.00"))
        self.assertEqual(self._sum(details["invoice"], "invoice_amount"),
                         current["invoice_matched"])
        self.assertEqual(self._sum(details["same_month"], "invoice_amount"),
                         current["same_month"])
        self.assertEqual(self._sum(details["cross"], "invoice_amount"), current["cross"])
        self.assertEqual(self._sum(details["month_accrual"], "amount"),
                         current["month_accrual"])
        self.assertEqual(current["invoice_matched"], Decimal("440.00"))
        self.assertEqual(current["cross"], Decimal("440.00"))
        self.assertEqual(current["same_month"], Decimal("0.00"))
        self.assertEqual(current["month_accrual"], Decimal("0.00"))
        # 两期之和 ≠ 本期（正是徽标曾经取错的地方）：期初 0 vs 400、本月暂估 400 vs 0
        self.assertEqual(total["opening_accrual"], Decimal("0.00"))
        self.assertEqual(total["month_accrual"], Decimal("400.00"))
        # 明细块内部自洽：当月到货到票 ＋ 本月回冲 ＝ 本月到票
        self.assertEqual(
            self._sum(details["same_month"], "invoice_amount")
            + self._sum(details["cross"], "invoice_amount"),
            self._sum(details["invoice"], "invoice_amount"),
        )
        # 徽标显示的是本期数字
        html = resp.content.decode("utf-8")
        self.assertEqual(self._pill(html, "期初暂估明细"), "400.00")
        self.assertEqual(self._pill(html, "本月到票明细"), "440.00")
        self.assertEqual(self._pill(html, "本月回冲明细"), "440.00")
        self.assertEqual(self._pill(html, "本月暂估明细"), "0.00")

    def test_first_period_opening_detail_comes_from_ledger(self):
        """首个期间没有上一期：期初暂估明细取自期初台账，徽标/合计同样是本期口径。"""
        resp = self.client.get(reverse("purchase_accruals", args=[self.june.id]))
        details = resp.context["details"]
        current = [r for r in resp.context["flow_rows"] if r["current"]][0]
        total = resp.context["flow_total"]
        self.assertIsNone(details["opening_from_previous"])
        self.assertEqual(self._sum(details["opening"], "amount"),
                         current["opening_accrual"])
        self.assertEqual(current["opening_accrual"], Decimal("0.00"))
        self.assertEqual(current["same_month"], Decimal("720.00"))
        self.assertEqual(current["month_accrual"], Decimal("400.00"))
        self.assertEqual(current["cross"], Decimal("0.00"))
        self.assertEqual(self._sum(details["month_accrual"], "amount"),
                         Decimal("400.00"))
        self.assertEqual(self._sum(details["cross"], "invoice_amount"),
                         Decimal("0.00"))
        # 回冲徽标必须是本期 0（跨期间合计是 440 —— 修好之前这里会显示 440）
        self.assertEqual(total["cross"], Decimal("440.00"))
        html = resp.content.decode("utf-8")
        self.assertEqual(self._pill(html, "本月回冲明细"), "0.00")
        self.assertEqual(self._pill(html, "本月到票明细"), "720.00")

class PurchaseNavTests(SettlementFixtureMixin, TestCase):
    """导航（2026-09-19 用户要求）：
    ①「采购匹配」改名为「采购发票」；
    ②新增「采购合同」入口（页面本来就有，只是没挂进导航）。
    同时锁住：期间首页流程卡片步骤 02 与导航同名同指向，且发票数按「张」计。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="nav-user", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 8, 1))

    def _nav_html(self, url_name="period_detail"):
        html = self.client.get(
            reverse(url_name, args=[self.period.id])
        ).content.decode("utf-8")
        # 只取「资料与采购 / 成本核算 …」这排期间导航，避免命中页面正文里的同名文字
        return html.split('class="period-links"')[1].split("</nav>")[0]

    def test_nav_label_is_purchase_invoice(self):
        """导航里叫「采购发票」，不再叫「采购匹配」。"""
        nav = self._nav_html()
        self.assertIn(
            'href="%s">采购发票</a>' % reverse("purchase_matches", args=[self.period.id]),
            nav,
        )
        self.assertNotIn("采购匹配", nav)

    def test_nav_has_purchase_contracts_entry(self):
        """导航新增「采购合同」，与「采购发票」同在「资料与采购」组且排在它后面。"""
        nav = self._nav_html()
        self.assertIn(
            'href="%s">采购合同</a>' % reverse("purchase_contracts", args=[self.period.id]),
            nav,
        )
        group = nav.split('data-nav="import"')[1].split("</div>")[0]
        self.assertIn(">采购发票</a>", group)
        self.assertIn(">采购合同</a>", group)
        self.assertLess(group.index(">采购发票</a>"), group.index(">采购合同</a>"))

    def test_nav_entries_are_reachable(self):
        """导航里的两个采购入口，点过去都能打开（200）且标题正确。"""
        for name, heading in (("purchase_matches", "采购发票匹配"),
                              ("purchase_contracts", "采购合同")):
            resp = self.client.get(reverse(name, args=[self.period.id]))
            self.assertEqual(resp.status_code, 200, name)
            self.assertIn("<h1>%s</h1>" % heading, resp.content.decode("utf-8"), name)

    def test_period_workflow_step_02_follows_nav(self):
        """期间首页流程卡片步骤 02：名称与导航一致，小字不再是合同条数。"""
        html = self.client.get(
            reverse("period_detail", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertIn("<strong>采购发票</strong>", html)
        self.assertNotIn("<strong>采购匹配</strong>", html)
        self.assertIn("张发票</small>", html)
        self.assertNotIn("条合同</small>", html)

    def test_period_workflow_counts_invoices_not_lines(self):
        """卡片显示的是「张」（不同发票号），不是明细行数 —— 与采购发票页口径一致。"""
        batch = self.create_batch(
            self.period, ImportBatch.Kind.PURCHASE_INVOICE, "nav-invoice"
        )
        # 同一张发票 2 行 + 另一张 1 行 → 张数 2、行数 3
        for i, invoice_no in enumerate(("INV-NAV-1", "INV-NAV-1", "INV-NAV-2")):
            PurchaseInvoiceLine.objects.create(
                period=self.period, batch=batch, invoice_no=invoice_no,
                invoice_date=date(2026, 8, 20), invoice_period=date(2026, 8, 1),
                material_code="RM-NAV-%d" % i, material_name="导航测试料",
                quantity=Decimal("10"), remaining_quantity=Decimal("10"),
                unit_price_excl_tax=Decimal("1"), amount_excl_tax=Decimal("10"),
            )
        resp = self.client.get(reverse("period_detail", args=[self.period.id]))
        html = resp.content.decode("utf-8")
        self.assertIn("2张发票</small>", html)
        self.assertNotIn("3张发票</small>", html)
        # 视图里两个口径并存：张 = 不同发票号，行 = 明细行数
        self.assertEqual(resp.context["counts"]["purchase_invoice_count"], 2)
        self.assertEqual(resp.context["counts"]["purchase_invoices"], 3)

    def test_purchase_invoice_page_browser_title(self):
        """采购发票页的浏览器标签标题与页内标题一致（原来还是旧名）。"""
        html = self.client.get(
            reverse("purchase_matches", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertIn("<title>采购发票匹配 ·", html)
        self.assertNotIn("<title>采购匹配 ·", html)
class ReceiptPeriodTests(SettlementFixtureMixin, TestCase):
    """「收货期间」（用户 2026-09-19 规则）。

    · 发票模板最左新增该列：**填了就以它决定这笔到票算「当月到货到票」还是「本月回冲」**；
      留空则沿用原规则（核销到的入库流水所属月份）。
    · 合同模板同样新增该列，但**只记录与展示，不参与计算**。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="receipt-period-user", password="test-password"
        )
        self.client.force_login(self.user)
        self.june = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        self.july = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.inventory_batch = self.create_batch(
            self.june, ImportBatch.Kind.INVENTORY, "inventory-rp"
        )
        self.contract_batch = self.create_batch(
            self.june, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-rp"
        )
        self.june_invoice_batch = self.create_batch(
            self.june, ImportBatch.Kind.PURCHASE_INVOICE, "invoice-rp-june"
        )
        self.july_invoice_batch = self.create_batch(
            self.july, ImportBatch.Kind.PURCHASE_INVOICE, "invoice-rp-july"
        )
        self.material = Material.objects.create(
            code="RM-RP", name="收货期间原料", uom="kg"
        )
        self.contract = PurchaseContractLine.objects.create(
            period=self.june, batch=self.contract_batch, contract_no="PC-RP",
            contract_date=date(2026, 6, 1), supplier_code="SUP-RP",
            supplier_name="收货期间供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("100"),
            remaining_quantity=Decimal("100"), unit_price_excl_tax=Decimal("10"),
        )
        # 6 月：入库 100@10 + 当月发票 60@12（当月到货到票 720），余 40 未到票（6 月暂估 400）
        self.june_invoice = PurchaseInvoiceLine.objects.create(
            period=self.june, batch=self.june_invoice_batch, invoice_no="PI-RP-06",
            invoice_date=date(2026, 6, 25), invoice_period=date(2026, 6, 1),
            contract_no="PC-RP", supplier_code="SUP-RP",
            supplier_name="收货期间供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("60"),
            remaining_quantity=Decimal("60"), unit_price_excl_tax=Decimal("12"),
            amount_excl_tax=Decimal("720"),
        )
        self.receipt = self.create_txn(
            self.june, self.inventory_batch, txn_date=date(2026, 6, 5),
            document_no="GR-RP", business_type="采购入库",
            direction=InventoryTransaction.Direction.IN, material=self.material,
            quantity=Decimal("100"), supplier_code="SUP-RP",
            supplier_name="收货期间供应商", contract_no="PC-RP",
        )
        auto_match_all(self.june)
        self.june.status = AccountingPeriod.Status.LOCKED
        self.june.save(update_fields=["status"])
        # 7 月：发票 40@11 核销 6 月入库 —— 默认应算「本月回冲 440」
        self.july_invoice = PurchaseInvoiceLine.objects.create(
            period=self.july, batch=self.july_invoice_batch, invoice_no="PI-RP-07",
            invoice_date=date(2026, 7, 20), invoice_period=date(2026, 7, 1),
            contract_no="PC-RP", supplier_code="SUP-RP",
            supplier_name="收货期间供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("40"),
            remaining_quantity=Decimal("40"), unit_price_excl_tax=Decimal("11"),
            amount_excl_tax=Decimal("440"),
        )
        auto_match_all(self.july)

    def _flow(self, period):
        rows, total = views.purchase_accrual_flow_tables(period, None)
        return {str(r["period"]): r for r in rows}

    def test_default_uses_matched_receipt_month(self):
        """留空＝沿用原规则：7 月这张发票核销 6 月入库 → 本月回冲（不是到货到票）。"""
        row = self._flow(self.july)[str(self.july)]
        self.assertEqual(row["cross"], Decimal("440.00"))
        self.assertEqual(row["same_month"], Decimal("0.00"))
        diff = build_purchase_difference_rows(self.july)
        self.assertEqual(len(diff), 1)
        self.assertTrue(diff[0].cross_period)
        # 归属月份＝回落后的入库流水月份（发票未填收货期间）
        self.assertEqual(diff[0].attribution_month, date(2026, 6, 1))

    def test_invoice_receipt_period_wins(self):
        """发票填了「收货期间」＝本期 → 这笔到票算「当月到货到票」，不再算回冲。"""
        self.july_invoice.receipt_period = date(2026, 7, 1)
        self.july_invoice.save(update_fields=["receipt_period"])
        row = self._flow(self.july)[str(self.july)]
        self.assertEqual(row["same_month"], Decimal("440.00"))
        self.assertEqual(row["cross"], Decimal("0.00"))
        diff = build_purchase_difference_rows(self.july)
        self.assertEqual(len(diff), 1)
        self.assertFalse(diff[0].cross_period)
        self.assertEqual(diff[0].attribution_month, date(2026, 7, 1))

    def test_receipt_period_can_turn_same_month_into_reversal(self):
        """反向也成立：本期入库 + 本月到票，若发票填的收货期间早于本期 → 算回冲。"""
        self.june_invoice.receipt_period = date(2026, 5, 1)
        self.june_invoice.save(update_fields=["receipt_period"])
        row = self._flow(self.june)[str(self.june)]
        self.assertEqual(row["cross"], Decimal("720.00"))
        self.assertEqual(row["same_month"], Decimal("0.00"))
        diff = build_purchase_difference_rows(self.june)
        self.assertEqual(len(diff), 1)
        self.assertTrue(diff[0].cross_period)

    def test_page_shows_receipt_period_column(self):
        """两个页面最左都有「收货期间」列，且显示 YYYY-MM。"""
        self.july_invoice.receipt_period = date(2026, 7, 1)
        self.july_invoice.save(update_fields=["receipt_period"])
        self.contract.receipt_period = date(2026, 6, 1)
        self.contract.save(update_fields=["receipt_period"])

        invoice_html = self.client.get(
            reverse("purchase_matches", args=[self.july.id])
        ).content.decode("utf-8")
        self.assertIn("<th>收货期间</th><th>发票日期</th>", invoice_html)
        self.assertIn("2026-07", invoice_html)

        contract_html = self.client.get(
            reverse("purchase_contracts", args=[self.june.id])
        ).content.decode("utf-8")
        self.assertIn(
            "<th>收货期间</th>"
            '<th class="num">已签收数量</th><th>合同号</th>',
            contract_html,
        )
        self.assertIn("2026-06", contract_html)

    def test_contract_receipt_period_does_not_change_calculation(self):
        """合同侧只记录：填了「收货期间」也不改变暂估/匹配结果。"""
        before = [
            (r.document_no, r.outstanding_quantity, r.provisional_amount)
            for r in build_purchase_accrual_rows(self.june)
        ]
        self.contract.refresh_from_db()
        self.contract_remaining_before = self.contract.remaining_quantity
        self.contract.receipt_period = date(2026, 6, 1)
        self.contract.save(update_fields=["receipt_period"])
        after = [
            (r.document_no, r.outstanding_quantity, r.provisional_amount)
            for r in build_purchase_accrual_rows(self.june)
        ]
        self.assertEqual(before, after)
        # 合同数量/剩余数量与匹配结果不受影响（与设置前一致）
        self.contract.refresh_from_db()
        self.assertEqual(
            self.contract.remaining_quantity, self.contract_remaining_before
        )

    def test_sample_templates_import_receipt_period(self):
        """两个上传模板的第 1 列就是「收货期间」，导入要落到字段上。"""
        from pathlib import Path
        from django.conf import settings
        from django.core.files.uploadedfile import SimpleUploadedFile
        from django.test.utils import override_settings

        template_dir = Path(settings.BASE_DIR) / "sample_templates"
        period = AccountingPeriod.objects.create(month=date(2026, 9, 1))
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                for kind, filename, model in (
                    (ImportBatch.Kind.PURCHASE_CONTRACT, "采购合同.csv",
                     PurchaseContractLine),
                    (ImportBatch.Kind.PURCHASE_INVOICE, "采购发票.csv",
                     PurchaseInvoiceLine),
                ):
                    content = (template_dir / filename).read_bytes()
                    uploaded = SimpleUploadedFile(filename, content, "text/csv")
                    batch = create_import_batch(
                        period=period, kind=kind, uploaded_file=uploaded,
                        user=self.user,
                    )
                    process_batch(batch)
                    rows = list(model.objects.filter(period=period))
                    self.assertTrue(rows, filename)
                    # 模板里两行样例分别填了 2026-06 / 2026-07（发票模板 1 行填 2026-06）
                    expected = {date(2026, 6, 1), date(2026, 7, 1)}
                    got = {row.receipt_period for row in rows}
                    self.assertTrue(
                        got and got.issubset(expected),
                        "%s 的收货期间未落到字段：%s" % (filename, got),
                    )
class ReceiptPeriodVoucherTests(SettlementFixtureMixin, TestCase):
    """「收货期间」与采购暂估/回冲凭证（用户 2026-09-19 选 C）。

    凭证是按「入库流水所属月份」分块的：入库属以前月份 → 走"前期暂估+本月到票"块，
    天然会**冲回以前月份挂的暂估**。所以即使「收货期间」把这张发票归到本期
    （报表口径显示"当月到货到票"），凭证仍然会把 6 月挂的暂估冲掉 —— 不会留下挂账。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="rp-voucher-user", password="test-password"
        )
        self.june = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        self.july = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.inventory_batch = self.create_batch(
            self.june, ImportBatch.Kind.INVENTORY, "inventory-rpv"
        )
        self.june_invoice_batch = self.create_batch(
            self.june, ImportBatch.Kind.PURCHASE_INVOICE, "invoice-rpv-june"
        )
        self.july_invoice_batch = self.create_batch(
            self.july, ImportBatch.Kind.PURCHASE_INVOICE, "invoice-rpv-july"
        )
        self.material = Material.objects.create(
            code="RM-RPV", name="凭证收货期间原料", uom="kg"
        )
        PurchaseContractLine.objects.create(
            period=self.june,
            batch=self.create_batch(
                self.june, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-rpv"
            ),
            contract_no="PC-RPV", contract_date=date(2026, 6, 1),
            supplier_code="SUP-RPV", supplier_name="凭证供应商",
            material_code=self.material.code, material_name=self.material.name,
            quantity=Decimal("100"), remaining_quantity=Decimal("100"),
            unit_price_excl_tax=Decimal("10"),
        )
        # 6 月：入库 100@10；发票 60@12 核销本月入库；余 40 未到票 → 6 月凭证挂暂估 400
        PurchaseInvoiceLine.objects.create(
            period=self.june, batch=self.june_invoice_batch, invoice_no="PI-RPV-06",
            invoice_date=date(2026, 6, 25), invoice_period=date(2026, 6, 1),
            contract_no="PC-RPV", supplier_code="SUP-RPV",
            supplier_name="凭证供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("60"),
            remaining_quantity=Decimal("60"), unit_price_excl_tax=Decimal("12"),
            amount_excl_tax=Decimal("720"),
        )
        self.receipt = self.create_txn(
            self.june, self.inventory_batch, txn_date=date(2026, 6, 5),
            document_no="GR-RPV", business_type="采购入库",
            direction=InventoryTransaction.Direction.IN, material=self.material,
            quantity=Decimal("100"), supplier_code="SUP-RPV",
            supplier_name="凭证供应商", contract_no="PC-RPV",
        )
        auto_match_all(self.june)
        # 7 月：发票 40@11 核销 6 月入库（默认＝回冲）
        self.july_invoice = PurchaseInvoiceLine.objects.create(
            period=self.july, batch=self.july_invoice_batch, invoice_no="PI-RPV-07",
            invoice_date=date(2026, 7, 20), invoice_period=date(2026, 7, 1),
            contract_no="PC-RPV", supplier_code="SUP-RPV",
            supplier_name="凭证供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("40"),
            remaining_quantity=Decimal("40"), unit_price_excl_tax=Decimal("11"),
            amount_excl_tax=Decimal("440"),
        )
        auto_match_all(self.july)

    def test_june_voucher_accrues_the_unmatched_receipt(self):
        """6 月凭证：未到票的 40 挂暂估（40 × 10＝400）。"""
        voucher = build_purchase_voucher(self.june)
        self.assertEqual(voucher["summary"]["accrual_amount"], Decimal("400.00"))
        self.assertEqual(voucher["summary"]["reversal_amount"], Decimal("0.00"))

    def test_voucher_reverses_prior_accrual_even_when_receipt_period_pulls_it_into_month(self):
        """收货期间＝本期（报表算"当月到货到票"）时，凭证仍冲回 6 月挂的暂估。"""
        self.july_invoice.receipt_period = date(2026, 7, 1)
        self.july_invoice.save(update_fields=["receipt_period"])

        # 报表口径：归本期 → 当月到货到票
        rows, _ = views.purchase_accrual_flow_tables(self.july, None)
        row = {str(r["period"]): r for r in rows}[str(self.july)]
        self.assertEqual(row["same_month"], Decimal("440.00"))
        self.assertEqual(row["cross"], Decimal("0.00"))

        # 凭证口径：仍走「前期暂估+本月到票」，把 6 月的暂估冲回（不留挂账）
        voucher = build_purchase_voucher(self.july)
        self.assertEqual(voucher["summary"]["reversal_amount"], Decimal("400.00"))
        self.assertEqual(
            voucher["summary"]["reversal_invoice_amount"], Decimal("440.00")
        )
        self.assertEqual(voucher["summary"]["same_month_amount"], Decimal("0.00"))
        self.assertEqual(voucher["summary"]["accrual_amount"], Decimal("0.00"))
        # 借贷平衡
        self.assertTrue(voucher["totals"]["balanced"])
        # 回冲明细里有这张发票
        self.assertIn(
            "PI-RPV-07",
            [r["invoice_no"] for r in voucher["reversal_rows"]],
        )
class ContractReceiptPeriodTests(SettlementFixtureMixin, TestCase):
    """合同的「收货期间」决定暂估归属（用户 2026-09-19 规则）。

    「合同里有到货签收的，就按照合同的收货期间来暂估，不要按照流水来」：
      · 合同行填了「收货期间」且该行有到货签收（已签收数量 > 0，或模板未提供该列）
        → 这份合同对应的入库，**暂估归属到合同的收货期间**；
      · 否则回落到入库流水月份（原行为）。
    影响：本月暂估 / 期末暂估余额 / 暂估凭证 / 到票归属（回冲与否）。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="contract-rp-user", password="test-password"
        )
        self.client.force_login(self.user)
        self.june = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        self.july = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.inventory_batch = self.create_batch(
            self.june, ImportBatch.Kind.INVENTORY, "inventory-crp"
        )
        self.contract_batch = self.create_batch(
            self.june, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-crp"
        )
        self.material = Material.objects.create(
            code="RM-CRP", name="合同收货期间原料", uom="kg"
        )
        self.contract = PurchaseContractLine.objects.create(
            period=self.june, batch=self.contract_batch, contract_no="PC-CRP",
            contract_date=date(2026, 6, 1), supplier_code="SUP-CRP",
            supplier_name="合同供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("100"),
            remaining_quantity=Decimal("100"), unit_price_excl_tax=Decimal("10"),
            received_quantity=Decimal("100"),
            receipt_period=date(2026, 7, 1),
        )
        # 6 月入库 100@10（未到票）—— 合同收货期间是 7 月，暂估应归 7 月
        self.receipt = self.create_txn(
            self.june, self.inventory_batch, txn_date=date(2026, 6, 5),
            document_no="GR-CRP", business_type="采购入库",
            direction=InventoryTransaction.Direction.IN, material=self.material,
            quantity=Decimal("100"), supplier_code="SUP-CRP",
            supplier_name="合同供应商", contract_no="PC-CRP",
        )
        auto_match_all(self.june)

    def _flow(self, period):
        rows, total = views.purchase_accrual_flow_tables(period, None)
        return {str(r["period"]): r for r in rows}

    def test_accrual_follows_contract_receipt_period(self):
        """合同收货期间＝7 月 → 6 月不挂暂估，7 月挂「本月暂估 1000」。"""
        june_rows = build_purchase_accrual_rows(self.june)
        self.assertEqual(june_rows, [])

        july_rows = build_purchase_accrual_rows(self.july)
        self.assertEqual(len(july_rows), 1)
        row = july_rows[0]
        self.assertEqual(row.receipt_month, date(2026, 7, 1))
        self.assertEqual(row.contract_month, date(2026, 7, 1))
        self.assertEqual(row.flow_month, date(2026, 6, 1))
        self.assertEqual(row.provisional_amount, Decimal("1000.00"))

        june = self._flow(self.june)[str(self.june)]
        july = self._flow(self.july)[str(self.july)]
        self.assertEqual(june["month_accrual"], Decimal("0.00"))
        self.assertEqual(june["ending_accrual"], Decimal("0.00"))
        self.assertEqual(july["opening_accrual"], Decimal("0.00"))
        self.assertEqual(july["month_accrual"], Decimal("1000.00"))
        self.assertEqual(july["ending_accrual"], Decimal("1000.00"))

    def test_zero_signed_quantity_falls_back_to_flow_month(self):
        """合同行没有到货签收（已签收数量 0）→ 不用它的收货期间，回到流水月份。"""
        self.contract.received_quantity = Decimal("0")
        self.contract.save(update_fields=["received_quantity"])
        rows = build_purchase_accrual_rows(self.june)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].receipt_month, date(2026, 6, 1))
        self.assertIsNone(rows[0].contract_month)
        june = self._flow(self.june)[str(self.june)]
        self.assertEqual(june["month_accrual"], Decimal("1000.00"))
        july = self._flow(self.july)[str(self.july)]
        self.assertEqual(july["month_accrual"], Decimal("0.00"))

    def test_missing_receipt_period_falls_back_to_flow_month(self):
        """合同没填收货期间 → 完全保持原行为。"""
        self.contract.receipt_period = None
        self.contract.save(update_fields=["receipt_period"])
        rows = build_purchase_accrual_rows(self.june)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].receipt_month, date(2026, 6, 1))

    def test_voucher_follows_contract_receipt_period(self):
        """凭证也跟着走：6 月不挂暂估，7 月凭证据挂（且入库金额按合同价 1000）。"""
        june_voucher = build_purchase_voucher(self.june)
        self.assertEqual(june_voucher["summary"]["accrual_amount"], Decimal("0.00"))
        july_voucher = build_purchase_voucher(self.july)
        self.assertEqual(july_voucher["summary"]["accrual_amount"], Decimal("1000.00"))
        self.assertTrue(july_voucher["totals"]["balanced"])

    def test_invoice_in_attribution_month_is_same_month(self):
        """7 月到票（发票未填收货期间）→ 按该入库的暂估归属（7 月）算「当月到货到票」。"""
        PurchaseInvoiceLine.objects.create(
            period=self.july,
            batch=self.create_batch(
                self.july, ImportBatch.Kind.PURCHASE_INVOICE, "invoice-crp"
            ),
            invoice_no="PI-CRP-07", invoice_date=date(2026, 7, 20),
            invoice_period=date(2026, 7, 1), contract_no="PC-CRP",
            supplier_code="SUP-CRP", supplier_name="合同供应商",
            material_code=self.material.code, material_name=self.material.name,
            quantity=Decimal("100"), remaining_quantity=Decimal("100"),
            unit_price_excl_tax=Decimal("11"), amount_excl_tax=Decimal("1100"),
        )
        auto_match_all(self.july)
        july = self._flow(self.july)[str(self.july)]
        self.assertEqual(july["same_month"], Decimal("1100.00"))
        self.assertEqual(july["cross"], Decimal("0.00"))
        # 全部到票 → 7 月不再有暂估
        self.assertEqual(july["month_accrual"], Decimal("0.00"))
        self.assertEqual(july["ending_accrual"], Decimal("0.00"))
        diff = build_purchase_difference_rows(self.july)
        self.assertEqual(len(diff), 1)
        self.assertFalse(diff[0].cross_period)

    def test_accrual_page_marks_contract_attribution(self):
        """月末暂估页：归属来自合同的行会标注「按合同收货期间」。"""
        html = self.client.get(
            reverse("purchase_accruals", args=[self.july.id])
        ).content.decode("utf-8")
        self.assertIn("按合同收货期间", html)
        self.assertIn("流水 2026-06", html)

    def test_template_received_quantity_is_stored(self):
        """上传模板里的「已签收数量」要存下来（判定"有到货签收"的依据）。"""
        from pathlib import Path
        from django.conf import settings
        from django.core.files.uploadedfile import SimpleUploadedFile
        from django.test.utils import override_settings

        template_dir = Path(settings.BASE_DIR) / "sample_templates"
        period = AccountingPeriod.objects.create(month=date(2026, 9, 1))
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                content = (template_dir / "采购合同.csv").read_bytes()
                uploaded = SimpleUploadedFile("采购合同.csv", content, "text/csv")
                batch = create_import_batch(
                    period=period, kind=ImportBatch.Kind.PURCHASE_CONTRACT,
                    uploaded_file=uploaded, user=self.user,
                )
                process_batch(batch)
        rows = list(PurchaseContractLine.objects.filter(period=period))
        self.assertTrue(rows)
        self.assertTrue(any(r.received_quantity and r.received_quantity > 0 for r in rows))
class AccrualDetailExportTests(SettlementFixtureMixin, TestCase):
    """月末暂估页 5 张明细小表各自一个「导出Excel」，按钮紧挨各自的「打印本表」（用户 2026-09-19）。

    · 按钮在标题右侧 .heading-actions 里，「打印本表」由 base.html 注入到紧跟其后的 .print-here；
    · 导出口径与页面「本期」一致：每张表导出文件的「合计」＝页面对应徽标（本期那一行的数字）。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="accrual-export-user", password="test-password"
        )
        self.client.force_login(self.user)
        self.june = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        self.july = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.inventory_batch = self.create_batch(
            self.june, ImportBatch.Kind.INVENTORY, "inventory-accrual-export"
        )
        self.july_inventory_batch = self.create_batch(
            self.july, ImportBatch.Kind.INVENTORY, "inventory-accrual-export-july"
        )
        self.contract_batch = self.create_batch(
            self.june, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-accrual-export"
        )
        self.june_invoice_batch = self.create_batch(
            self.june, ImportBatch.Kind.PURCHASE_INVOICE, "invoice-accrual-export-june"
        )
        self.july_invoice_batch = self.create_batch(
            self.july, ImportBatch.Kind.PURCHASE_INVOICE, "invoice-accrual-export-july"
        )
        self.material = Material.objects.create(
            code="RM-AEX", name="导出原料", uom="kg"
        )
        PurchaseContractLine.objects.create(
            period=self.june, batch=self.contract_batch, contract_no="PC-AEX",
            contract_date=date(2026, 6, 1), supplier_code="SUP-AEX",
            supplier_name="导出供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("200"),
            remaining_quantity=Decimal("200"), unit_price_excl_tax=Decimal("10"),
        )
        # 6 月：入库 100@10 + 发票 60@12 → 当到货到票 720；余 40 → 6 月暂估 400
        PurchaseInvoiceLine.objects.create(
            period=self.june, batch=self.june_invoice_batch, invoice_no="PI-AEX-06",
            invoice_date=date(2026, 6, 25), invoice_period=date(2026, 6, 1),
            contract_no="PC-AEX", supplier_code="SUP-AEX",
            supplier_name="导出供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("60"),
            remaining_quantity=Decimal("60"), unit_price_excl_tax=Decimal("12"),
            amount_excl_tax=Decimal("720"),
        )
        self.create_txn(
            self.june, self.inventory_batch, txn_date=date(2026, 6, 5),
            document_no="GR-AEX", business_type="采购入库",
            direction=InventoryTransaction.Direction.IN, material=self.material,
            quantity=Decimal("100"), supplier_code="SUP-AEX",
            supplier_name="导出供应商", contract_no="PC-AEX",
        )
        auto_match_all(self.june)
        self.june.status = AccountingPeriod.Status.LOCKED
        self.june.save(update_fields=["status"])
        # 7 月：① 发票 40@11 核销 6 月入库 → 本月回冲 440
        #       ② 入库 50@10 + 发票 30@10 → 当月到货到票 300，余 20 → 本月暂估 200
        PurchaseInvoiceLine.objects.create(
            period=self.july, batch=self.july_invoice_batch, invoice_no="PI-AEX-07A",
            invoice_date=date(2026, 7, 20), invoice_period=date(2026, 7, 1),
            contract_no="PC-AEX", supplier_code="SUP-AEX",
            supplier_name="导出供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("40"),
            remaining_quantity=Decimal("40"), unit_price_excl_tax=Decimal("11"),
            amount_excl_tax=Decimal("440"),
        )
        self.create_txn(
            self.july, self.july_inventory_batch, txn_date=date(2026, 7, 10),
            document_no="GR-AEX-07", business_type="采购入库",
            direction=InventoryTransaction.Direction.IN, material=self.material,
            quantity=Decimal("50"), supplier_code="SUP-AEX",
            supplier_name="导出供应商", contract_no="PC-AEX",
        )
        auto_match_all(self.july)
        PurchaseInvoiceLine.objects.create(
            period=self.july, batch=self.july_invoice_batch, invoice_no="PI-AEX-07B",
            invoice_date=date(2026, 7, 25), invoice_period=date(2026, 7, 1),
            contract_no="PC-AEX", supplier_code="SUP-AEX",
            supplier_name="导出供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("30"),
            remaining_quantity=Decimal("30"), unit_price_excl_tax=Decimal("10"),
            amount_excl_tax=Decimal("300"),
        )
        auto_match_all(self.july)
        self.url = reverse("purchase_accruals", args=[self.july.id])

    # ---- 小工具 ----
    def _current(self):
        resp = self.client.get(self.url)
        return [r for r in resp.context["flow_rows"] if r["current"]][0]

    def _sheet(self, kind, period=None):
        resp = self.client.get(
            reverse("export_listing", args=[(period or self.july).id, kind])
        )
        self.assertEqual(resp.status_code, 200, kind)
        workbook = load_workbook(BytesIO(resp.content), read_only=True)
        rows = list(workbook.worksheets[0].iter_rows(values_only=True))
        return workbook.worksheets[0].title, rows

    def _total_row(self, rows):
        body = [r for r in rows[1:] if r and any(v not in (None, "") for v in r)]
        self.assertTrue(body, "导出没有数据行")
        self.assertEqual(body[-1][0], "合计", "最后一行应是合计行")
        return body[-1]

    def _dec(self, value):
        return Decimal(str(value if value is not None else 0)).quantize(Decimal("0.01"))

    # ---- 用例 ----
    def test_each_detail_block_has_export_button_next_to_print(self):
        """5 张明细表都有「导出Excel」，且紧跟 .print-here（打印本表落点）。

        用户 2026-09-19：有筛选栏的 4 张表，把「金额徽标 / 导出Excel / 打印本表」挪进筛选栏，
        排在「筛选 / 清空」旁边；只有期初暂估明细（无筛选栏）还留在标题右侧。
        """
        html = self.client.get(self.url).content.decode("utf-8")
        for heading, kind in (
            ("期初暂估明细", "purchase_accrual_opening"),
            ("本月到票明细", "purchase_accrual_invoice"),
            ("当月到货到票明细", "purchase_accrual_same_month"),
            ("本月暂估明细", "purchase_accrual_month"),
            ("本月回冲明细", "purchase_accrual_cross"),
        ):
            seg = html.split("<h2>%s</h2></div>" % heading, 1)[1].split("</section>", 1)[0]
            link = reverse("export_listing", args=[self.july.id, kind])
            self.assertIn(link, seg, heading)
            self.assertIn('class="print-here"', seg, heading)
            self.assertLess(seg.index(link), seg.index('class="print-here"'), heading)
            self.assertIn("导出Excel", seg, heading)
            if heading == "期初暂估明细":
                # 没有筛选栏：金额徽标 + 导出按钮仍排在标题右侧
                head_block = seg.split('class="detail-filter', 1)[0]
                self.assertIn('class="heading-actions"', head_block, heading)
                self.assertIn("导出Excel", head_block, heading)
            else:
                # 有筛选栏：标题右侧不再放按钮，金额徽标 / 导出Excel / 打印本表都在筛选栏那一排
                bar = seg.split('class="filter-actions"', 1)[1].split("</div>", 1)[0]
                self.assertNotIn("heading-actions", seg.split('class="detail-filter', 1)[0], heading)
                self.assertIn(">筛选</button>", bar, heading)
                self.assertIn(">清空</a>", bar, heading)
                self.assertIn("detail-filter-pill", bar, heading)
                self.assertIn("导出Excel", bar, heading)
                self.assertLess(bar.index(">清空</a>"), bar.index("detail-filter-pill"), heading)
                self.assertLess(bar.index("detail-filter-pill"), bar.index("导出Excel"), heading)
                self.assertLess(bar.index("导出Excel"), bar.index('class="print-here"'), heading)

    def test_detail_filter_pill_shows_filtered_total(self):
        """筛选栏里的金额徽标＝当前表格里的合计：未筛选＝本期全口径，筛选后＝命中金额。"""
        html = self.client.get(self.url).content.decode("utf-8")
        m = re.search(r'detail-filter-pill">本月到票 ([\d,]+\.\d\d)<', html)
        self.assertIsNotNone(m)
        self.assertEqual(self._dec(m.group(1).replace(",", "")),
                         self._dec(self._current()["invoice_matched"]))
        # 筛「导出供应商」（本月到票 2 行：440 回冲 + 300 当到货到票）
        html = self.client.get(self.url, {"inv_supplier": "导出供应商"}).content.decode("utf-8")
        m = re.search(r'detail-filter-pill">本月到票 ([\d,]+\.\d\d)（筛选后）<', html)
        self.assertIsNotNone(m, "筛选后徽标应带「（筛选后）」")
        self.assertEqual(self._dec(m.group(1).replace(",", "")), Decimal("740.00"))
        # 上表那一行仍是本期全口径，不受明细表筛选影响
        resp = self.client.get(self.url, {"inv_supplier": "导出供应商"})
        self.assertEqual(self._dec(resp.context["flow_current"]["invoice_matched"]),
                         self._dec(self._current()["invoice_matched"]))

    def test_opening_export_matches_page_pill(self):
        """期初暂估明细：合计＝上表「期初暂估」（＝6 月期末 400）。"""
        title, rows = self._sheet("purchase_accrual_opening")
        self.assertEqual(title, "期初暂估明细")
        self.assertEqual(rows[0][:3], ("归属月份", "入库日期", "入库单号"))
        self.assertEqual(self._dec(self._total_row(rows)[10]),
                         self._dec(self._current()["opening_accrual"]))

    def test_month_export_matches_page_pill(self):
        """本月暂估明细：合计＝上表「本月暂估」（7 月入库 50、到票 30 → 余 20×10＝200）。"""
        title, rows = self._sheet("purchase_accrual_month")
        self.assertEqual(title, "本月暂估明细")
        total = self._total_row(rows)
        self.assertEqual(self._dec(total[10]), self._dec(self._current()["month_accrual"]))
        self.assertEqual(self._dec(total[10]), Decimal("200.00"))

    def test_same_month_and_cross_export_match_page_pills(self):
        """当月到货到票 300、本月回冲 440（发票金额口径），与上表两列一致。"""
        _, same_rows = self._sheet("purchase_accrual_same_month")
        same_total = self._total_row(same_rows)
        # 用户 2026-09-19：不再出「发票号」，金额列索引随之前移（发票金额＝索引 9）
        self.assertEqual(self._dec(same_total[9]), Decimal("300.00"))
        self.assertEqual(self._dec(same_total[9]),
                         self._dec(self._current()["same_month"]))

        _, cross_rows = self._sheet("purchase_accrual_cross")
        cross_total = self._total_row(cross_rows)
        # 第 8 列（索引 7）是「冲回暂估金额」，末列备注写发票金额合计
        self.assertIn("发票金额合计 440.00", cross_total[-1])
        # 第 8 列（索引 7）是「冲回暂估金额」（暂估口径 400）
        self.assertEqual(self._dec(cross_total[7]), Decimal("400.00"))
        resp = self.client.get(self.url)
        expected_provisional = sum(
            (r.provisional_amount or Decimal("0")
             for r in resp.context["details"]["cross"]),
            Decimal("0"),
        )
        self.assertEqual(self._dec(cross_total[7]), self._dec(expected_provisional))

    def test_invoice_export_equals_same_month_plus_cross(self):
        """本月到票明细导出：合计＝当月到货到票＋本月回冲，且逐行标注期间性质。"""
        title, rows = self._sheet("purchase_accrual_invoice")
        self.assertEqual(title, "本月到票明细")
        total = self._total_row(rows)
        current = self._current()
        self.assertEqual(self._dec(total[9]), self._dec(current["invoice_matched"]))
        self.assertEqual(
            self._dec(total[9]),
            self._dec(current["same_month"]) + self._dec(current["cross"]),
        )
        kinds = {r[11] for r in rows[1:] if r and r[0] not in (None, "", "合计")}
        self.assertIn("本月回冲", kinds)
        self.assertIn("当月到货到票", kinds)

    def test_invoice_no_column_removed_from_page_and_export(self):
        """本月到票 / 当月到货到票 / 本月回冲 三张表不再有「发票号」列（用户 2026-09-19）。"""
        html = self.client.get(self.url).content.decode("utf-8")
        for heading in ("本月到票明细", "当月到货到票明细", "本月回冲明细"):
            seg = html.split("<h2>%s</h2></div>" % heading, 1)[1].split("</section>", 1)[0]
            self.assertNotIn("<th>发票号</th>", seg, heading)
        for kind in ("purchase_accrual_invoice", "purchase_accrual_same_month",
                     "purchase_accrual_cross"):
            _, rows = self._sheet(kind)
            self.assertNotIn("发票号", rows[0], kind)

    def test_detail_filters_narrow_rows_and_keep_flow_numbers(self):
        """4 张明细表各自的筛选：只影响这几张表，上表数字不变、导出也跟着筛。"""
        # 本月到票明细：按供应商筛（数据里只有「导出供应商」，命中全部 2 行）
        resp = self.client.get(self.url, {"inv_supplier": "导出供应商"})
        self.assertEqual(len(resp.context["details"]["invoice"]), 2)
        self.assertEqual(resp.context["flow_current"]["invoice_matched"],
                         self._current()["invoice_matched"])
        # 换个不存在的供应商 → 0 行
        resp = self.client.get(self.url, {"inv_supplier": "不存在的供应商"})
        self.assertEqual(resp.context["details"]["invoice"], [])
        # 按物料编码筛
        resp = self.client.get(self.url, {"inv_material": "RM-AEX"})
        self.assertEqual(len(resp.context["details"]["invoice"]), 2)
        # 本月暂估明细：归属月份＝7 月（有）、＝6 月（无，6 月的行不在「本月暂估」里）
        resp = self.client.get(self.url, {"ma_month": "2026-07"})
        self.assertEqual(len(resp.context["details"]["month_accrual"]), 1)
        resp = self.client.get(self.url, {"ma_month": "2026-06"})
        self.assertEqual(resp.context["details"]["month_accrual"], [])
        # 本月回冲明细：按供应商筛（1 行）
        # ⚠️ 供应商列显示的是「名称优先」，所以筛选也要按名称（编码只在没有名称时才显示）
        resp = self.client.get(self.url, {"cr_supplier": "导出供应商"})
        self.assertEqual(len(resp.context["details"]["cross"]), 1)
        resp = self.client.get(self.url, {"cr_supplier": "不存在的供应商"})
        self.assertEqual(resp.context["details"]["cross"], [])

    def test_detail_filter_carries_over_to_export_link(self):
        """「导出Excel」带上本表筛选条件（筛选后导出＝所见），且其它表的导出不会被带上。"""
        html = self.client.get(self.url, {"inv_supplier": "SUP-AEX"}).content.decode("utf-8")
        seg = html.split("<h2>本月到票明细</h2></div>", 1)[1].split("</section>", 1)[0]
        self.assertIn("inv_supplier=SUP-AEX", seg)
        # 本月暂估明细：导出链接只带它自己的条件（inv_* 只是作为隐藏域帮用户保留条件，
        # 不会写进它的导出链接）
        seg2 = html.split("<h2>本月暂估明细</h2></div>", 1)[1].split("</section>", 1)[0]
        export_link = [part for part in seg2.split('href="') if "export-list" in part[:80]][0]
        self.assertNotIn("inv_supplier", export_link.split('"')[0])
        # 隐藏域仍带着本表以外的条件，避免提交本表筛选时把别的筛选冲掉
        self.assertIn('name="inv_supplier"', seg2)
        # 导出文件也按同样的条件过滤：
        # ① 筛「不存在的供应商」→ 没有数据行（只留表头，也就不写合计行）
        resp = self.client.get(
            reverse("export_listing", args=[self.july.id, "purchase_accrual_invoice"]),
            {"inv_supplier": "不存在的供应商"},
        )
        workbook = load_workbook(BytesIO(resp.content), read_only=True)
        rows = list(workbook.worksheets[0].iter_rows(values_only=True))
        body = [r for r in rows[1:] if r and any(v not in (None, "") for v in r)]
        self.assertEqual(body, [])
        # ② 筛「导出供应商」→ 2 行明细 + 1 行合计，合计＝本月到票（＝当月到货到票＋本月回冲）
        resp = self.client.get(
            reverse("export_listing", args=[self.july.id, "purchase_accrual_invoice"]),
            {"inv_supplier": "导出供应商"},
        )
        workbook = load_workbook(BytesIO(resp.content), read_only=True)
        rows = list(workbook.worksheets[0].iter_rows(values_only=True))
        body = [r for r in rows[1:] if r and any(v not in (None, "") for v in r)]
        self.assertEqual(len(body), 3)  # 2 行明细 + 合计
        self.assertEqual(body[-1][0], "合计")
        self.assertEqual(self._dec(body[-1][9]),
                         self._dec(self._current()["invoice_matched"]))

    def test_exports_are_current_period_only(self):
        """明细块是「本期」口径：6 月页导出的是 6 月的明细。"""
        _, rows = self._sheet("purchase_accrual_same_month", period=self.june)
        total = self._total_row(rows)
        self.assertEqual(self._dec(total[9]), Decimal("720.00"))
        resp = self.client.get(reverse("purchase_accruals", args=[self.june.id]))
        june_row = [r for r in resp.context["flow_rows"] if r["current"]][0]
        self.assertEqual(self._dec(total[9]), self._dec(june_row["same_month"]))
class WorkshopSettleGuardTests(SettlementFixtureMixin, TestCase):
    """车间结转的前置校验：已结转 / 前序未完成 要「立刻」给出友好提示，不重算、不留失败记录。

    背景（用户 2026-09-19 反馈「针织结转不了」）：针织车间其实已经在 13:36 结转完成，
    但在它还没跑完时页面上仍显示「结转本车间」，用户又点了一次：
    旧代码先花约 2 分钟做 `_prepare_period`（重算承接 + 自动匹配 + 回填），
    之后才检查状态，于是白等两分钟后弹「车间结转失败：针织车间已经完成结转。」，
    还在结算记录里留下一条 status=failed 的批次。现在改成：先校验、后重算，
    并且这类用户操作问题按提示（info）返回、不写失败记录。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="settle-guard-user", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 8, 1))
        self.settled = WorkshopSettlement.objects.create(
            period=self.period, workshop="ZhenZhi", sequence=2,
            status=WorkshopSettlement.Status.SETTLED, converged=True,
            iterations=1, order_cost=Decimal("100.00"),
            production_receipt_amount=Decimal("80.00"),
        )
        self.pending = WorkshopSettlement.objects.create(
            period=self.period, workshop="JiZhi", sequence=3,
        )
        self.first = WorkshopSettlement.objects.create(
            period=self.period, workshop="RanZheng", sequence=1,
        )

    def _patch_prepare(self):
        """把耗时的 _prepare_period 换成「被调用就记录」的桩，用来证明没有白等重算。"""
        from settlement.services import costing as costing_module

        calls = []
        original = costing_module._prepare_period
        costing_module._prepare_period = lambda period: calls.append(period) or {}
        self.addCleanup(
            setattr, costing_module, "_prepare_period", original
        )
        return calls

    def test_settled_workshop_is_rejected_without_recompute(self):
        """已结转的车间：立刻报「已经结转完成」，不重算、不写失败批次。"""
        from settlement.services import costing as costing_module

        calls = self._patch_prepare()
        with self.assertRaisesMessage(
            costing_module.SettlePreconditionError, "已经结转完成"
        ):
            costing_module.settle_workshop(
                self.period, self.settled.id, self.user
            )
        self.assertEqual(calls, [], "不该为重算白等（_prepare_period 被调用了）")
        self.assertFalse(
            SettlementRun.objects.filter(period=self.period).exists(),
            "这类提示不该留下结算批次（尤其不该是 failed）",
        )
        # 车间结果没被动过
        self.settled.refresh_from_db()
        self.assertEqual(self.settled.status, WorkshopSettlement.Status.SETTLED)
        self.assertEqual(self.settled.order_cost, Decimal("100.00"))

    def test_predecessor_pending_is_rejected_without_recompute(self):
        """前序车间没结转：用中文车间名提示，且不跑重算、不留批次。"""
        from settlement.services import costing as costing_module

        calls = self._patch_prepare()
        with self.assertRaisesMessage(
            costing_module.SettlePreconditionError, "请先完成前序车间：染色车间"
        ):
            costing_module.settle_workshop(
                self.period, self.pending.id, self.user
            )
        self.assertEqual(calls, [])
        self.assertFalse(SettlementRun.objects.filter(period=self.period).exists())

    def test_view_returns_hint_not_failure(self):
        """页面点「结转本车间」点到一个已结转的车间：给提示，不给红字失败。"""
        url = reverse("settle_workshop", args=[self.period.id, self.settled.id])
        resp = self.client.post(url, follow=True)
        self.assertEqual(resp.status_code, 200)
        messages = list(resp.context["messages"])
        texts = [str(m) for m in messages]
        self.assertTrue(
            any("已经结转完成" in text for text in texts),
            "应提示已经结转完成，实际：%s" % texts,
        )
        self.assertNotIn("error", {m.level_tag for m in messages})
        self.assertFalse(
            SettlementRun.objects.filter(
                period=self.period, status=SettlementRun.Status.FAILED
            ).exists()
        )

    def test_running_run_records_workshop_name(self):
        """抢占结转时就把车间名写进批次：运行中的横幅能显示是哪个车间在结转。"""
        from settlement.services import costing as costing_module

        run = costing_module._claim_settle_run(
            self.period, mode="workshop", workshop="ZhenZhi"
        )
        self.assertEqual(run.summary.get("workshop"), "ZhenZhi")
        context = views._running_settle_context(self.period)
        self.assertIsNotNone(context)
        self.assertEqual(context["workshop"], "针织车间")
        # 冲突提示里也应带车间名
        with self.assertRaisesMessage(ValueError, "针织车间的结转还在进行中"):
            costing_module._claim_settle_run(
                self.period, mode="workshop", workshop="JiZhi"
            )
        run.status = SettlementRun.Status.SUCCESS
        run.finished_at = timezone.now()
        run.save(update_fields=["status", "finished_at"])

    def test_settle_button_stays_disabled_while_running(self):
        """提交后按钮一直禁用（以前 3 秒后放开，用户会重复点击）。"""
        html = self.client.get(
            reverse("workshop_settlements", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertIn("btn.disabled = true", html)
        self.assertNotIn("btn.disabled = false", html)
class ContractMatchModeTests(SettlementFixtureMixin, TestCase):
    """采购合同的「处理方式」（用户 2026-09-19 选①：给合同表加 match_mode 字段，与发票页一致）。

    口径：
      · 自动匹配：只有「自动匹配」的合同参与（忽略/手工的都不自动占用剩余数量）；
      · 忽略匹配：不计入「合同金额」，也不作为暂估价格兜底，且不往后续期间自动承接；
      · 手工匹配：只保留人工指定的关系（与采购发票同款行为）。
    背景：合同页此前就有「处理方式」列与筛选，但合同模型上没有该字段 ——
    列一直显示空白、点筛选会 500（FieldError）。本用例锁住「能显示、能筛选、能修改」。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="contract-mode-user", password="test-password"
        )
        self.client.force_login(self.user)
        self.june = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        self.july = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.contract_batch = self.create_batch(
            self.june, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-mode"
        )
        self.inventory_batch = self.create_batch(
            self.june, ImportBatch.Kind.INVENTORY, "inventory-mode"
        )
        self.material = Material.objects.create(
            code="RM-MODE", name="处理方式原料", uom="kg"
        )
        self.contract = PurchaseContractLine.objects.create(
            period=self.june, batch=self.contract_batch, contract_no="PC-MODE",
            contract_date=date(2026, 6, 1), supplier_code="SUP-MODE",
            supplier_name="处理方式供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("100"),
            remaining_quantity=Decimal("100"), unit_price_excl_tax=Decimal("10"),
        )
        self.receipt = self.create_txn(
            self.june, self.inventory_batch, txn_date=date(2026, 6, 5),
            document_no="GR-MODE", business_type="采购入库",
            direction=InventoryTransaction.Direction.IN, material=self.material,
            quantity=Decimal("100"), supplier_code="SUP-MODE",
            supplier_name="处理方式供应商", contract_no="PC-MODE",
        )

    def _allocations(self):
        return MatchAllocation.objects.filter(
            kind=MatchAllocation.Kind.PURCHASE_CONTRACT,
            contract_line=self.contract,
        )

    def test_auto_match_uses_auto_contract(self):
        """默认（自动匹配）照旧：入库被合同占用 → 合同暂估单价可用。"""
        auto_match_all(self.june)
        self.assertEqual(self._allocations().count(), 1)
        self.contract.refresh_from_db()
        self.assertEqual(self.contract.remaining_quantity, Decimal("0.000000"))

    def test_ignored_contract_is_not_auto_matched(self):
        """忽略匹配 → 不再自动占用入库，剩余数量保持合同数量。"""
        set_source_match_mode(
            self.june, self.contract, MatchingMode.IGNORED, "重复导入"
        )
        self.contract.refresh_from_db()
        self.assertEqual(self.contract.match_mode, MatchingMode.IGNORED)
        self.assertEqual(self.contract.ignore_reason, "重复导入")
        self.assertEqual(self._allocations().count(), 0)
        self.assertEqual(self.contract.remaining_quantity, Decimal("100.000000"))
        # 再跑一次自动匹配也不会重新占用
        auto_match_all(self.june)
        self.assertEqual(self._allocations().count(), 0)

    def test_ignored_contract_excluded_from_amount_and_price_fallback(self):
        """忽略匹配 → 不计入「合同金额」，也不作为暂估价格兜底。"""
        auto_match_all(self.june)
        self.assertEqual(
            views.contract_amount_by_month().get(self.june.month), Decimal("1000.00")
        )
        set_source_match_mode(self.june, self.contract, MatchingMode.IGNORED)
        self.assertEqual(
            views.contract_amount_by_month().get(self.june.month, Decimal("0.00")),
            Decimal("0.00"),
        )
        # 入库没有合同匹配、也没有原始金额 → 单价来源不应该是合同价兜底
        basis = provisional_basis(self.receipt)
        self.assertNotEqual(basis["source"], "合同暂估")
        self.assertNotEqual(basis["source"], "合同暂估（价格参考）")

    def test_manual_contract_keeps_manual_allocation_only(self):
        """手工匹配：合同转「手工匹配」，只保留人工关系（自动匹配不再补一条）。"""
        manual_match(
            self.june,
            MatchAllocation.Kind.PURCHASE_CONTRACT,
            self.receipt.id,
            self.contract.id,
            Decimal("40"),
        )
        self.contract.refresh_from_db()
        self.assertEqual(self.contract.match_mode, MatchingMode.MANUAL)
        allocations = self._allocations()
        self.assertEqual(allocations.count(), 1)
        self.assertEqual(allocations[0].matched_quantity, Decimal("40.000000"))
        self.assertTrue(allocations[0].is_manual)
        # 手工匹配后可改回自动
        set_source_match_mode(self.june, self.contract, MatchingMode.AUTO)
        self.contract.refresh_from_db()
        self.assertEqual(self.contract.match_mode, MatchingMode.AUTO)
        self.assertEqual(self._allocations().count(), 1)

    def test_contract_page_shows_mode_and_edit_link(self):
        """合同页：「处理方式」列有值、有「打开修改」入口（原来该列一直空白）。"""
        set_source_match_mode(
            self.june, self.contract, MatchingMode.IGNORED, "作废合同"
        )
        html = self.client.get(
            reverse("purchase_contracts", args=[self.june.id])
        ).content.decode("utf-8")
        self.assertIn("忽略匹配", html)
        self.assertIn("作废合同", html)
        self.assertIn(
            reverse(
                "source_document_detail",
                args=[self.june.id, "purchase_contract", self.contract.id],
            ),
            html,
        )

    def test_contract_page_filter_by_match_mode(self):
        """「处理方式」筛选可用（此前引用不存在的字段，会 500）。"""
        set_source_match_mode(self.june, self.contract, MatchingMode.IGNORED)
        other = PurchaseContractLine.objects.create(
            period=self.june, batch=self.contract_batch, contract_no="PC-MODE-2",
            contract_date=date(2026, 6, 2), supplier_code="SUP-MODE",
            supplier_name="处理方式供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("50"),
            remaining_quantity=Decimal("50"), unit_price_excl_tax=Decimal("12"),
        )
        resp = self.client.get(
            reverse("purchase_contracts", args=[self.june.id]),
            {"match_mode": MatchingMode.IGNORED},
        )
        self.assertEqual(resp.status_code, 200)
        ids = [row.id for row in resp.context["page"].object_list]
        self.assertIn(self.contract.id, ids)
        self.assertNotIn(other.id, ids)

    def test_contract_detail_page_can_set_and_reset_mode(self):
        """合同详情页（打开修改）：POST 能改处理方式，改回自动后自动匹配恢复。"""
        url = reverse(
            "source_document_detail",
            args=[self.june.id, "purchase_contract", self.contract.id],
        )
        html = self.client.get(url).content.decode("utf-8")
        self.assertIn("采购合同匹配明细", html)
        self.assertIn("PC-MODE", html)

        resp = self.client.post(
            url,
            {"action": "set_mode", "match_mode": MatchingMode.IGNORED,
             "ignore_reason": "重复导入"},
        )
        self.assertEqual(resp.status_code, 302)
        self.contract.refresh_from_db()
        self.assertEqual(self.contract.match_mode, MatchingMode.IGNORED)
        self.assertEqual(self._allocations().count(), 0)

        self.client.post(
            url,
            {"action": "set_mode", "match_mode": MatchingMode.AUTO,
             "ignore_reason": ""},
        )
        self.contract.refresh_from_db()
        self.assertEqual(self.contract.match_mode, MatchingMode.AUTO)
        self.assertEqual(self._allocations().count(), 1)

    def test_contract_detail_page_can_add_manual_match(self):
        """合同详情页：能给这份合同手工指定入库流水。"""
        url = reverse(
            "source_document_detail",
            args=[self.june.id, "purchase_contract", self.contract.id],
        )
        resp = self.client.post(
            url,
            {"action": "add_match", "transaction_id": self.receipt.id,
             "matched_quantity": "30"},
        )
        self.assertEqual(resp.status_code, 302)
        self.contract.refresh_from_db()
        self.assertEqual(self.contract.match_mode, MatchingMode.MANUAL)
        self.assertEqual(self._allocations().count(), 1)
        self.assertEqual(
            self._allocations()[0].matched_quantity, Decimal("30.000000")
        )

    def test_ignored_contract_is_not_carried_forward(self):
        """忽略匹配的合同不往后续期间承接（否则每月都会被带出来）。"""
        from settlement.services.rollforward import _carry_contracts
        from settlement.models import MatchingMode as _MM

        keep = PurchaseContractLine.objects.create(
            period=self.june, batch=self.contract_batch, contract_no="PC-KEEP",
            contract_date=date(2026, 6, 3), supplier_code="SUP-MODE",
            supplier_name="处理方式供应商", material_code="RM-KEEP",
            material_name="另一种原料", quantity=Decimal("20"),
            remaining_quantity=Decimal("20"), unit_price_excl_tax=Decimal("9"),
        )
        set_source_match_mode(self.june, self.contract, _MM.IGNORED, "作废")
        _carry_contracts(self.july, self.june)
        carried = PurchaseContractLine.objects.filter(period=self.july)
        self.assertEqual([row.contract_no for row in carried], ["PC-KEEP"])
        self.assertEqual(carried[0].carried_from_id, keep.id)

    def test_transaction_detail_omits_ignored_contracts(self):
        """库存流水的手工匹配候选里不出现被忽略的合同。"""
        set_source_match_mode(self.june, self.contract, MatchingMode.IGNORED)
        groups = views._transaction_source_groups(self.june, self.receipt)
        contract_group = [g for g in groups if g["label"] == "采购合同"][0]
        self.assertEqual(list(contract_group["items"]), [])
class ContractInvoiceDateTests(SettlementFixtureMixin, TestCase):
    """合同模板新增「收票日期」列（用户 2026-09-19）。

    位置：紧挨「收货期间」之后（模板第 2 列、页面第 2 列），「已签收数量」之前。
    含义（为以后「到票」也能从合同表取数做准备）：
      · 留空      → 尚未收票（仍在暂估）
      · ＝收货期间 → 当月到货当月到票
      · ＞收货期间 → 跨期到票（要冲回以前月份挂的暂估）
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="invoice-date-user", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 9, 1))

    def _import_template(self, filename="采购合同.csv"):
        from pathlib import Path

        from django.conf import settings
        from django.core.files.uploadedfile import SimpleUploadedFile
        from django.test.utils import override_settings

        template_dir = Path(settings.BASE_DIR) / "sample_templates"
        content = (template_dir / filename).read_bytes()
        uploaded = SimpleUploadedFile(filename, content, "text/csv")
        period = AccountingPeriod.objects.create(month=date(2026, 10, 1))
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                batch = create_import_batch(
                    period=period,
                    kind=ImportBatch.Kind.PURCHASE_CONTRACT,
                    uploaded_file=uploaded,
                    user=self.user,
                )
                process_batch(batch)
        return list(PurchaseContractLine.objects.filter(period=period))

    def test_sample_template_has_invoice_date_between_receipt_period_and_received(self):
        """示例模板：第 1 列收货期间、第 2 列收票日期、已签收数量在其后。"""
        from pathlib import Path

        from django.conf import settings

        path = Path(settings.BASE_DIR) / "sample_templates" / "采购合同.csv"
        text = path.read_bytes().decode("utf-8-sig")
        header = text.splitlines()[0].split(",")
        self.assertEqual(header[0], "收货期间")
        self.assertEqual(header[1], "收票日期")
        self.assertLess(header.index("收票日期"), header.index("已签收数量"))

    def test_import_fills_invoice_date(self):
        """导入时「收票日期」落到字段：月份写法（2026-07）与完整日期都支持。"""
        rows = self._import_template()
        self.assertTrue(rows)
        got = {row.invoice_date for row in rows}
        self.assertTrue(
            got.issubset({None, date(2026, 7, 1)}),
            "收货期间列不应被当成收票日期；实际：%s" % got,
        )
        self.assertIn(date(2026, 7, 1), got, "样例第 2 行的收票日期应落成 2026-07-01")

    def test_import_accepts_full_date(self):
        """填完整日期（2026-09-15）时按原样存，不强行归到月初。"""
        from settlement.services.imports import _import_purchase_contract

        batch = self.create_batch(
            self.period, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-invdate3"
        )
        _import_purchase_contract(
            batch,
            {
                "contract_no": "PC-ID-1",
                "material_code": "RM-ID",
                "quantity": "10",
                "unit_price_excl_tax": "5",
                "receipt_period": "2026-07",
                "invoice_date": "2026-09-15",
                "received_quantity": "10",
            },
        )
        row = PurchaseContractLine.objects.get(contract_no="PC-ID-1")
        self.assertEqual(row.receipt_period, date(2026, 7, 1))
        self.assertEqual(row.invoice_date, date(2026, 9, 15))

    def test_import_without_invoice_date_stays_empty(self):
        """没填「收票日期」→ 存空（含义＝尚未收票，仍挂暂估）。"""
        from settlement.services.imports import _import_purchase_contract

        batch = self.create_batch(
            self.period, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-invdate4"
        )
        _import_purchase_contract(
            batch,
            {
                "contract_no": "PC-ID-4",
                "material_code": "RM-ID4",
                "quantity": "7",
                "unit_price_excl_tax": "3",
                "receipt_period": "2026-09",
                "received_quantity": "7",
            },
        )
        row = PurchaseContractLine.objects.get(contract_no="PC-ID-4")
        self.assertIsNone(row.invoice_date)
        self.assertEqual(row.receipt_period, date(2026, 9, 1))

    def test_page_and_export_show_invoice_date_column(self):
        """页面与导出：收货期间 | 已签收数量 依次排列，且带「金额」列（用户 2026-09-19 调整）。"""
        batch = self.create_batch(
            self.period, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-invdate"
        )
        PurchaseContractLine.objects.create(
            period=self.period, batch=batch, contract_no="PC-ID-2",
            contract_date=date(2026, 9, 1), supplier_code="SUP-ID",
            supplier_name="收票日期供应商", material_code="RM-ID2",
            material_name="收票日期原料", quantity=Decimal("100"),
            remaining_quantity=Decimal("100"), unit_price_excl_tax=Decimal("10"),
            receipt_period=date(2026, 9, 1), invoice_date=date(2026, 9, 20),
            received_quantity=Decimal("100"),
        )
        html = self.client.get(
            reverse("purchase_contracts", args=[self.period.id])
        ).content.decode("utf-8")
        # 页面上第一张表现在是「采购合同汇总」，这里只看「合同明细」那张表
        detail_html = html.split("<h2>合同明细</h2>", 1)[1]
        block = detail_html.split("</thead>")[0]
        heads = re.findall(r"<th[^>]*>([^<]+)</th>", block)
        self.assertEqual(heads[:3], ["收货期间", "已签收数量", "合同号"])
        self.assertIn("金额", heads)
        self.assertNotIn("收票日期", heads)
        self.assertIn("2026-09", html)

        resp = self.client.get(
            reverse("export_listing", args=[self.period.id, "purchase_contract"])
        )
        workbook = load_workbook(BytesIO(resp.content), read_only=True)
        rows = list(workbook.worksheets[0].iter_rows(values_only=True))
        self.assertEqual(rows[0][:3], ("收货期间", "已签收数量", "合同号"))
        self.assertIn("金额", rows[0], "导出也要带「金额」列")

    def test_carry_forward_preserves_invoice_date(self):
        """承接要带上「收票日期」，否则后续期间看不到收票信息。"""
        from settlement.services.rollforward import _carry_contracts

        batch = self.create_batch(
            self.period, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-invdate2"
        )
        source = PurchaseContractLine.objects.create(
            period=self.period, batch=batch, contract_no="PC-ID-3",
            contract_date=date(2026, 9, 1), supplier_code="SUP-ID",
            supplier_name="收票日期供应商", material_code="RM-ID3",
            material_name="收票日期原料", quantity=Decimal("50"),
            remaining_quantity=Decimal("50"), unit_price_excl_tax=Decimal("8"),
            receipt_period=date(2026, 9, 1), invoice_date=date(2026, 10, 8),
            received_quantity=Decimal("50"),
        )
        next_period = AccountingPeriod.objects.create(month=date(2026, 11, 1))
        _carry_contracts(next_period, self.period)
        carried = PurchaseContractLine.objects.get(
            period=next_period, carried_from=source
        )
        self.assertEqual(carried.invoice_date, date(2026, 10, 8))
        self.assertEqual(carried.receipt_period, date(2026, 9, 1))
class InvoiceAmountBasisTests(SettlementFixtureMixin, TestCase):
    """发票核销金额以「不含税金额」列为准（用户 2026-09-19）。

    背景：发票文件里「不含税单价 × 数量」与「不含税金额」在部分行不一致
    （原币/税口径混用），按单价乘数量会把核销金额放大。
    现在：核销金额 = 发票行不含税金额 × 核销数量 ÷ 发票数量，
    且一张发票分多笔核销时合计正好等于该发票金额（尾差归最后一次分摊）。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="inv-basis-user", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        self.inventory_batch = self.create_batch(
            self.period, ImportBatch.Kind.INVENTORY, "inventory-invbasis"
        )
        self.invoice_batch = self.create_batch(
            self.period, ImportBatch.Kind.PURCHASE_INVOICE, "invoice-invbasis"
        )
        self.material = Material.objects.create(
            code="RM-INVB", name="发票口径原料", uom="kg"
        )

    def _invoice(self, *, quantity, unit_price, amount, **kwargs):
        values = dict(
            period=self.period, batch=self.invoice_batch, invoice_no="PI-INVB",
            invoice_date=date(2026, 6, 20), invoice_period=date(2026, 6, 1),
            supplier_code="SUP-INVB", supplier_name="发票口径供应商",
            material_code=self.material.code, material_name=self.material.name,
            quantity=quantity, remaining_quantity=quantity,
            unit_price_excl_tax=unit_price, amount_excl_tax=amount,
        )
        values.update(kwargs)
        return PurchaseInvoiceLine.objects.create(**values)

    def _receipt(self, quantity, document_no="GR-INVB"):
        return self.create_txn(
            self.period, self.inventory_batch, txn_date=date(2026, 6, 5),
            document_no=document_no, business_type="采购入库",
            direction=InventoryTransaction.Direction.IN, material=self.material,
            quantity=quantity, supplier_code="SUP-INVB",
            supplier_name="发票口径供应商",
        )

    def _allocations(self):
        return MatchAllocation.objects.filter(
            period=self.period, kind=MatchAllocation.Kind.PURCHASE_INVOICE
        )

    def test_matched_amount_follows_invoice_amount_column(self):
        """单价×数量=1000 但金额列=900 → 核销金额按 900 认。"""
        self._invoice(quantity=Decimal("100"), unit_price=Decimal("10"),
                      amount=Decimal("900"))
        self._receipt(Decimal("100"))
        auto_match_all(self.period)
        allocs = self._allocations()
        self.assertEqual(len(allocs), 1)
        self.assertEqual(allocs[0].amount, Decimal("900.00"))
        # 单位价格也按金额折算（900/100=9）
        self.assertEqual(Decimal(allocs[0].unit_price).quantize(Decimal("0.01")),
                         Decimal("9.00"))

    def test_split_allocations_sum_to_invoice_amount(self):
        """一张发票分两笔核销（40+60）时，核销金额合计＝发票金额，不留尾差。"""
        self._invoice(quantity=Decimal("100"), unit_price=Decimal("10"),
                      amount=Decimal("901.37"))
        self._receipt(Decimal("40"), document_no="GR-INVB-A")
        self._receipt(Decimal("60"), document_no="GR-INVB-B")
        auto_match_all(self.period)
        allocs = self._allocations()
        self.assertEqual(len(allocs), 2)
        self.assertEqual(
            sum((a.amount for a in allocs), Decimal("0")), Decimal("901.37")
        )

    def test_falls_back_to_unit_price_when_amount_missing(self):
        """发票没有金额列时，退回 数量 × 单价（老行为）。"""
        invoice = self._invoice(quantity=Decimal("100"), unit_price=Decimal("10"),
                                amount=None)
        # 导入外的场景：金额为空 → 核销金额 = 数量×单价
        self.assertIsNone(invoice.amount_excl_tax)
        self._receipt(Decimal("100"))
        auto_match_all(self.period)
        self.assertEqual(self._allocations()[0].amount, Decimal("1000.00"))

    def test_opening_accrual_respects_invoice_receipt_period(self):
        """核销期初台账时，发票自己的「收货期间」也要生效（以前被忽略 → 6,720 错分）。"""
        opening_batch = self.create_batch(
            self.period, ImportBatch.Kind.OPENING_PURCHASE_ACCRUAL, "opening-invb"
        )
        opening = OpeningPurchaseAccrual.objects.create(
            period=self.period, batch=opening_batch,
            original_period=date(2025, 12, 1),
            document_no="OPEN-INVB", document_date=date(2025, 12, 20),
            supplier_code="SUP-INVB", supplier_name="发票口径供应商",
            material_code=self.material.code,
            material_name=self.material.name, opening_quantity=Decimal("100"),
            original_quantity=Decimal("100"),
            remaining_quantity=Decimal("100"),
            provisional_unit_price=Decimal("10"),
            opening_amount=Decimal("1000"),
        )
        self._invoice(quantity=Decimal("100"), unit_price=Decimal("10"),
                      amount=Decimal("1000"),
                      receipt_period=self.period.month)
        auto_match_all(self.period)
        allocs = self._allocations()
        self.assertEqual(len(allocs), 1)
        self.assertIsNone(allocs[0].inventory_transaction)
        self.assertEqual(allocs[0].opening_purchase_accrual_id, opening.id)
        # 报表：归属本月 → 当月到货到票（而不是回冲）
        resp = self.client.get(
            reverse("purchase_accruals", args=[self.period.id])
        )
        row = [r for r in resp.context["flow_rows"] if r["current"]][0]
        self.assertEqual(row["same_month"], Decimal("1000.00"))
        self.assertEqual(row["cross"], Decimal("0.00"))


class ContractAmountBasisTests(SettlementFixtureMixin, TestCase):
    """「合同金额」列以合同文件的「原币未税总价（报表用）」为准（用户 2026-09-19）。"""

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="contract-amt-basis", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        self.batch = self.create_batch(
            self.period, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-amt-basis"
        )

    def _contract(self, **overrides):
        values = dict(
            period=self.period, batch=self.batch, contract_no="PC-AB",
            contract_date=date(2026, 6, 1), supplier_code="SUP-AB",
            supplier_name="合同金额供应商", material_code="RM-AB",
            material_name="合同金额原料", quantity=Decimal("100"),
            remaining_quantity=Decimal("100"), unit_price_excl_tax=Decimal("10"),
            receipt_period=self.period.month, received_quantity=Decimal("100"),
        )
        values.update(overrides)
        return PurchaseContractLine.objects.create(**values)

    def _flow_row(self):
        resp = self.client.get(reverse("purchase_accruals", args=[self.period.id]))
        return [r for r in resp.context["flow_rows"] if r["current"]][0]

    def test_uses_uploaded_total_amount(self):
        """文件给了未税总价 → 用它（86400），而不是 数量×单价（604800）。"""
        self._contract(quantity=Decimal("10800"),
                       unit_price_excl_tax=Decimal("56"),
                       amount_excl_tax=Decimal("86400"))
        self.assertEqual(self._flow_row()["contract_amount"], Decimal("86400.00"))

    def test_keeps_long_decimals_without_per_row_rounding(self):
        """多行长小数时按原值累加、只在最后取整（不逐行四舍五入）。"""
        self._contract(contract_no="PC-AB-1", amount_excl_tax=Decimal("1234.567890"))
        self._contract(contract_no="PC-AB-2", amount_excl_tax=Decimal("2345.678900"))
        self.assertEqual(self._flow_row()["contract_amount"], Decimal("3580.246790"))

    def test_falls_back_to_quantity_times_price(self):
        """没有未税总价列的老数据 → 退回 数量 × 单价。"""
        self._contract(amount_excl_tax=None)
        self.assertEqual(self._flow_row()["contract_amount"], Decimal("1000.00"))

    def test_import_maps_total_amount_column(self):
        """导入：识别「原币未税总价（报表用）」列并落到字段。"""
        from settlement.services.imports import _import_purchase_contract

        _import_purchase_contract(
            self.batch,
            {
                "contract_no": "PC-AB-IMP", "material_code": "RM-AB-IMP",
                "quantity": "10", "unit_price_excl_tax": "5",
                "amount_excl_tax": "48.67", "receipt_period": "2026-06",
            },
        )
        row = PurchaseContractLine.objects.get(contract_no="PC-AB-IMP")
        self.assertEqual(row.amount_excl_tax, Decimal("48.67"))
class PlanANoticeTests(SettlementFixtureMixin, TestCase):
    """方案A（用户 2026-09-19 选定）：合同表管到货/暂估，发票表管到票/实际价。

    这里锁两件事：
      ① 本期**没上传采购发票**时，月末暂估页要明确提示「到票/回冲按 0 计算」，
         免得把 0 误读成"本期没有到票"；上传发票后提示消失。
      ② 「收票日期」这一列**只记录展示、不参与计算**（暂估/到票/回冲仍按发票核销算）。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="plan-a-user", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        self.inventory_batch = self.create_batch(
            self.period, ImportBatch.Kind.INVENTORY, "inventory-plan-a"
        )
        self.material = Material.objects.create(
            code="RM-PLANA", name="方案A原料", uom="kg"
        )
        self.receipt = self.create_txn(
            self.period, self.inventory_batch, txn_date=date(2026, 6, 5),
            document_no="GR-PLANA", business_type="采购入库",
            direction=InventoryTransaction.Direction.IN, material=self.material,
            quantity=Decimal("100"), supplier_code="SUP-PLANA",
            supplier_name="方案A供应商",
        )
        self.url = reverse("purchase_accruals", args=[self.period.id])

    def test_notice_shown_when_no_invoice_uploaded(self):
        """没传发票 → 有提示，且明说按 0 计算。"""
        resp = self.client.get(self.url)
        html = resp.content.decode("utf-8")
        self.assertTrue(resp.context["no_invoice_data"])
        self.assertIn("没有上传采购发票", html)
        self.assertIn("message-warning", html)
        self.assertIn("按 <strong>0</strong> 计算", html)
        # 暂估本身照常算（入库 100 有原始金额兜底/缺价，但这行要在）
        self.assertTrue(resp.context["flow_rows"])

    def test_notice_hidden_after_invoice_uploaded(self):
        """传了发票 → 提示消失。"""
        batch = self.create_batch(
            self.period, ImportBatch.Kind.PURCHASE_INVOICE, "invoice-plan-a"
        )
        PurchaseInvoiceLine.objects.create(
            period=self.period, batch=batch, invoice_no="PI-PLANA",
            invoice_date=date(2026, 6, 25), invoice_period=date(2026, 6, 1),
            supplier_code="SUP-PLANA", supplier_name="方案A供应商",
            material_code=self.material.code, material_name=self.material.name,
            quantity=Decimal("60"), remaining_quantity=Decimal("60"),
            unit_price_excl_tax=Decimal("12"), amount_excl_tax=Decimal("720"),
        )
        resp = self.client.get(self.url)
        self.assertFalse(resp.context["no_invoice_data"])
        self.assertNotIn("没有上传采购发票", resp.content.decode("utf-8"))

    def test_invoice_date_is_record_only(self):
        """收票日期只记录：改它不影响暂估行、到票/回冲（仍按发票核销算）。"""
        batch = self.create_batch(
            self.period, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-plan-a"
        )
        contract = PurchaseContractLine.objects.create(
            period=self.period, batch=batch, contract_no="PC-PLANA",
            contract_date=date(2026, 6, 1), supplier_code="SUP-PLANA",
            supplier_name="方案A供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("100"),
            remaining_quantity=Decimal("100"), unit_price_excl_tax=Decimal("10"),
            receipt_period=date(2026, 6, 1), received_quantity=Decimal("100"),
            invoice_date=date(2026, 6, 20),
        )
        auto_match_all(self.period)

        def snapshot():
            resp = self.client.get(self.url)
            row = [r for r in resp.context["flow_rows"] if r["current"]][0]
            accruals = [
                (r.document_no, r.receipt_quantity, r.outstanding_quantity, r.provisional_amount)
                for r in build_purchase_accrual_rows(self.period)
            ]
            return (
                row["invoice_matched"], row["same_month"], row["cross"],
                row["month_accrual"], row["ending_accrual"], accruals,
            )

        before = snapshot()
        # 把收票日期挪到 8 月、甚至清空 —— 数字都不该变
        contract.invoice_date = date(2026, 8, 15)
        contract.save(update_fields=["invoice_date"])
        self.assertEqual(snapshot(), before)
        contract.invoice_date = None
        contract.save(update_fields=["invoice_date"])
        self.assertEqual(snapshot(), before)
        # 页面上不再展示「收票日期」（用户 2026-09-19 要求删掉这一列，连带页头那句说明也去掉了）
        html = self.client.get(
            reverse("purchase_contracts", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertNotIn("收票日期", html)
        self.assertIn("金额", html)
class ContractPageLayoutTests(SettlementFixtureMixin, TestCase):
    """采购合同页的版式（用户 2026-09-19 要求）：

    ① 删掉「合同号 / 供应商 / 品名编码 / 处理方式 + 筛选 / 清空」整个筛选模块；
    ② 「导出Excel」从页头挪到表格上方那一排，紧挨「打印本表」（打印在右）。
    （视图层的筛选能力仍保留：带 ?match_mode= 之类的链接照样能过滤——只是页面上不再给控件。）
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="contract-layout-user", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 9, 1))
        self.batch = self.create_batch(
            self.period, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-layout"
        )
        PurchaseContractLine.objects.create(
            period=self.period, batch=self.batch, contract_no="PC-LAY",
            contract_date=date(2026, 9, 1), supplier_code="SUP-LAY",
            supplier_name="版式供应商", material_code="RM-LAY",
            material_name="版式原料", quantity=Decimal("10"),
            remaining_quantity=Decimal("10"), unit_price_excl_tax=Decimal("5"),
        )
        self.url = reverse("purchase_contracts", args=[self.period.id])

    def test_filter_module_is_removed(self):
        """筛选模块（含筛选/清空按钮与四个筛选控件）已整块删除。"""
        html = self.client.get(self.url).content.decode("utf-8")
        # 只看「筛选模块那段标记」是否存在（base.html 的脚本里也会出现 filter-card 字样，不能用裸词判断）
        self.assertNotIn('<section class="section-card filter-card">', html)
        self.assertNotIn('class="filter-grid"', html)
        self.assertNotIn('name="contract_no"', html)
        self.assertNotIn('name="supplier"', html)
        self.assertNotIn('name="material"', html)
        self.assertNotIn('name="match_mode"', html)
        self.assertNotIn(">筛选</button>", html)
        self.assertNotIn(">清空</a>", html)
        # 页面上不该残留这些筛选提示文字（多行 {# #} 注释会漏成正文，这里一并防住）
        self.assertNotIn("包含关键字", html)

    def test_export_button_sits_next_to_print(self):
        """「导出Excel」在表格那一排，且紧挨 .print-here（打印本表注入点）。"""
        html = self.client.get(self.url).content.decode("utf-8")
        export_url = reverse("export_listing", args=[self.period.id, "purchase_contract"])
        self.assertIn(export_url, html)
        self.assertIn('class="print-here"', html)
        # 导出按钮必须排在 .print-here 之前（打印本表落在它右边）
        self.assertLess(html.index(export_url), html.index('class="print-here"'))
        # 页头不再放导出按钮（页头与表格两段分开看）
        page_heading = html.split("</section>", 1)[0]
        self.assertNotIn(export_url, page_heading)
        # 同一排：导出按钮与 .print-here 处在同一个 .heading-actions 容器里
        block = html[html.index('class="heading-actions"', html.index("<h2>合同明细</h2>")):]
        block = block[: block.index("</div>", block.index('class="print-here"'))]
        self.assertIn(export_url, block)
        self.assertIn('class="print-here"', block)

    def test_export_still_works_without_page_filters(self):
        """删掉筛选控件后，导出仍能正常出文件（且表头不变）。"""
        resp = self.client.get(
            reverse("export_listing", args=[self.period.id, "purchase_contract"])
        )
        self.assertEqual(resp.status_code, 200)
        workbook = load_workbook(BytesIO(resp.content), read_only=True)
        rows = list(workbook.worksheets[0].iter_rows(values_only=True))
        self.assertEqual(rows[0][:3], ("收货期间", "已签收数量", "合同号"))


class ContractSummaryTests(SettlementFixtureMixin, TestCase):
    """采购合同页「采购合同汇总」表 + 明细表清理（用户 2026-09-19 要求）。

    汇总表字段：收货期间 / 合同数量 / 签收数量 / 合同金额 / 签收金额，按「收货期间」分组 + 一行合计。
      · 合同金额 ＝ 合同行不含税总价（文件的「原币未税总价（报表用）」，老数据退回 数量×单价）
      · 签收金额 ＝ 已签收数量 × 不含税单价
    明细表：去掉「明细ID」小字；合同号**有就填、没有就留空**（不显示占位符）。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="contract-summary-user", password="test-password"
        )
        self.client.force_login(self.user)
        self.may = AccountingPeriod.objects.create(month=date(2026, 5, 1))
        self.june = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        self.batch = self.create_batch(
            self.may, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-summary"
        )

    def _contract(self, **overrides):
        values = dict(
            period=self.may, batch=self.batch, contract_no="PC-S",
            contract_date=date(2026, 5, 1), supplier_code="SUP-S",
            supplier_name="汇总供应商", material_code="RM-S",
            material_name="汇总原料", quantity=Decimal("100"),
            remaining_quantity=Decimal("100"), unit_price_excl_tax=Decimal("10"),
        )
        values.update(overrides)
        return PurchaseContractLine.objects.create(**values)

    def _page(self):
        resp = self.client.get(reverse("purchase_contracts", args=[self.may.id]))
        self.assertEqual(resp.status_code, 200)
        return resp.context

    def test_summary_groups_by_receipt_period(self):
        """按收货期间分组 + 合计：数量/金额逐组正确。"""
        self._contract(
            contract_no="PC-S1", receipt_period=date(2026, 5, 1),
            quantity=Decimal("100"), received_quantity=Decimal("80"),
            unit_price_excl_tax=Decimal("10"), amount_excl_tax=Decimal("1000"),
        )
        self._contract(
            contract_no="PC-S2", receipt_period=date(2026, 6, 1),
            quantity=Decimal("50"), received_quantity=Decimal("50"),
            unit_price_excl_tax=Decimal("20"), amount_excl_tax=Decimal("1000"),
        )
        context = self._page()
        rows = {r["receipt_period"]: r for r in context["summary_rows"]}
        self.assertEqual(set(rows), {date(2026, 5, 1), date(2026, 6, 1)})
        # 5 月：合同 100 / 签收 80 / 合同金额 1000（文件总价）/ 签收金额 80×10=800
        self.assertEqual(rows[date(2026, 5, 1)]["contract_quantity"], Decimal("100"))
        self.assertEqual(rows[date(2026, 5, 1)]["received_quantity"], Decimal("80"))
        self.assertEqual(rows[date(2026, 5, 1)]["contract_amount"], Decimal("1000"))
        self.assertEqual(rows[date(2026, 5, 1)]["received_amount"], Decimal("800.00"))
        # 6 月：50 / 50 / 1000 / 50×20=1000
        self.assertEqual(rows[date(2026, 6, 1)]["contract_amount"], Decimal("1000"))
        self.assertEqual(rows[date(2026, 6, 1)]["received_amount"], Decimal("1000.00"))
        total = context["summary_total"]
        self.assertEqual(total["contract_quantity"], Decimal("150"))
        self.assertEqual(total["received_quantity"], Decimal("130"))
        self.assertEqual(total["contract_amount"], Decimal("2000"))
        self.assertEqual(total["received_amount"], Decimal("1800.00"))

    def test_amounts_fall_back_when_file_has_no_total(self):
        """没有「原币未税总价」的老数据：合同金额退回 数量 × 单价；没有签收数量则签收金额为 0。"""
        self._contract(
            contract_no="PC-S3", receipt_period=date(2026, 5, 1),
            quantity=Decimal("30"), received_quantity=None,
            unit_price_excl_tax=Decimal("5"), amount_excl_tax=None,
        )
        row = [r for r in self._page()["summary_rows"]][0]
        self.assertEqual(row["contract_amount"], Decimal("150.00"))
        self.assertEqual(row["received_amount"], Decimal("0.00"))

    def test_contract_without_receipt_period_goes_to_last_group(self):
        """没填收货期间的合同单独成组、排在最后（页面上显示「未填」）。"""
        self._contract(
            contract_no="PC-S4", receipt_period=date(2026, 5, 1),
            quantity=Decimal("10"), received_quantity=Decimal("10"),
            unit_price_excl_tax=Decimal("2"), amount_excl_tax=Decimal("20"),
        )
        self._contract(
            contract_no="PC-S5", receipt_period=None,
            quantity=Decimal("7"), received_quantity=Decimal("7"),
            unit_price_excl_tax=Decimal("3"), amount_excl_tax=Decimal("21"),
        )
        context = self._page()
        self.assertEqual([r["receipt_period"] for r in context["summary_rows"]],
                         [date(2026, 5, 1), None])
        total = context["summary_total"]
        self.assertEqual(total["contract_amount"], Decimal("41"))
        self.assertEqual(total["received_amount"], Decimal("41.00"))

    def test_summary_renders_with_columns_and_total(self):
        """页面渲染：汇总表 5 列 + 合计行 + 口径说明。"""
        self._contract(
            contract_no="PC-S6", receipt_period=date(2026, 5, 1),
            quantity=Decimal("10"), received_quantity=Decimal("10"),
            unit_price_excl_tax=Decimal("2"), amount_excl_tax=Decimal("20"),
        )
        html = self.client.get(
            reverse("purchase_contracts", args=[self.may.id])
        ).content.decode("utf-8")
        self.assertIn("<h2>采购合同汇总</h2>", html)
        block = html.split('class="contract-summary-table"')[1].split("</table>")[0]
        self.assertEqual(
            re.findall(r"<th[^>]*>([^<]+)</th>", block.split("</thead>")[0]),
            ["收货期间", "合同数量", "签收数量", "合同金额", "签收金额"],
        )
        self.assertIn("合计", block)
        self.assertIn("签收金额</strong> ＝ 已签收数量 × 不含税单价", html)

    def test_detail_rows_drop_detail_id_and_blanks_contract_no(self):
        """明细表：没有「明细ID」字样；空合同号的行该格留空。"""
        self._contract(
            contract_no="", receipt_period=date(2026, 5, 1),
            quantity=Decimal("5"), received_quantity=Decimal("5"),
            unit_price_excl_tax=Decimal("1"), amount_excl_tax=Decimal("5"),
        )
        html = self.client.get(
            reverse("purchase_contracts", args=[self.may.id])
        ).content.decode("utf-8")
        self.assertNotIn("明细ID", html)
        # 明细表那一行：合同号单元格应为空（不带 — 占位符）；只看「合同明细」那张表
        detail_html = html.split("<h2>合同明细</h2>", 1)[1]
        body = detail_html.split("<tbody>")[1].split("</tbody>")[0]
        cells = re.findall(r"<td[^>]*>(.*?)</td>", body, re.S)
        contract_no_cell = re.sub(r"\s+", "", re.sub(r"<[^>]+>", "", cells[2]))
        self.assertEqual(contract_no_cell, "")
class PulpEndingInventoryTests(SettlementFixtureMixin, TestCase):
    """浆粕（浆粕 / 微纤 / JP-）也进「期末存货」并按台账结转（用户 2026-09-19）。

    背景：以前 `_persist_ending_by_batch` 把浆粕成品**整类排除**在通用期末表之外
    （注释写"浆粕成品期末独立核算"），结果是浆粕没有期末存货行 →
    下期承接（`_carry_inventory` 优先读上期期末存货）拿不到它、库存断档。

    现在浆粕与普通物料同规则（"就按照台账来"）：
      期末数量 = 期初（浆粕期初上传）+ 本月入库 − 本月出库（浆粕流水上传）
      期末金额 = 数量 × 月末加权价（来自浆粕成本核算）
    若另外上传了浆粕期末，仍以人工为准。
    """

    PULP_CODE = "21.03"
    PULP_NAME = "芳纶浆粕DFP-103"
    JEAN_CODE = "22.01"
    JEAN_NAME = "芳纶微纤AMF-100"

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="pulp-ending-user", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.opening_batch = self.create_batch(
            self.period, ImportBatch.Kind.PULP_OPENING, "pulp-opening"
        )
        self.ledger_batch = self.create_batch(
            self.period, ImportBatch.Kind.PULP_LEDGER, "pulp-ledger"
        )
        self.material = Material.objects.create(
            code=self.PULP_CODE, name=self.PULP_NAME, uom="kg"
        )
        # 浆粕期初成品（单独上传的台账）
        OpeningInventoryBalance.objects.create(
            period=self.period,
            batch=self.opening_batch,
            material_code=self.PULP_CODE,
            material_name=self.PULP_NAME,
            quantity=Decimal("1000"),
            amount=Decimal("60000"),
            unit_price=Decimal("60"),
            is_system_generated=False,
        )
        # 浆粕流水（生产入库 + 销售出库）
        self.create_txn(
            self.period, self.ledger_batch, txn_date=date(2026, 7, 10),
            document_no="PULP-IN", business_type="生产入库",
            direction=InventoryTransaction.Direction.IN,
            material=self.material, quantity=Decimal("500"),
        )
        self.create_txn(
            self.period, self.ledger_batch, txn_date=date(2026, 7, 20),
            document_no="PULP-OUT", business_type="销售出库",
            direction=InventoryTransaction.Direction.OUT,
            material=self.material, quantity=Decimal("200"),
        )
        self.summaries = {
            self.PULP_CODE: {
                "material_name": self.PULP_NAME,
                "weighted_average_price": Decimal("62"),
            }
        }

    def _system_ending(self):
        return EndingInventoryBalance.objects.filter(
            period=self.period, is_system_generated=True
        )

    def test_pulp_gets_system_ending_from_ledger(self):
        """浆粕按台账算出期末（数量＝期初+入库−出库，金额＝数量×加权价），不再被排除。"""
        from settlement.services.costing import _persist_ending_by_batch

        _persist_ending_by_batch(self.period, self.summaries)
        row = self._system_ending().get(material_code=self.PULP_CODE)
        self.assertEqual(row.quantity, Decimal("1300.000000"))   # 1000 + 500 − 200
        self.assertEqual(row.amount, Decimal("80600.00"))        # 1300 × 62
        self.assertEqual(row.unit_price, Decimal("62"))
        self.assertEqual(row.ending_date, date(2026, 7, 31))

    def test_jean_pulp_named_material_not_excluded(self):
        """名称含「微纤」的物料同样进期末（以前也被排除）。"""
        from settlement.services.costing import _persist_ending_by_batch

        material = Material.objects.create(
            code=self.JEAN_CODE, name=self.JEAN_NAME, uom="kg"
        )
        OpeningInventoryBalance.objects.create(
            period=self.period, batch=self.opening_batch,
            material_code=self.JEAN_CODE, material_name=self.JEAN_NAME,
            quantity=Decimal("100"), amount=Decimal("5700"),
            unit_price=Decimal("57"), is_system_generated=False,
        )
        summaries = dict(self.summaries)
        summaries[self.JEAN_CODE] = {
            "material_name": self.JEAN_NAME,
            "weighted_average_price": Decimal("57"),
        }
        _persist_ending_by_batch(self.period, summaries)
        row = self._system_ending().get(material_code=self.JEAN_CODE)
        self.assertEqual(row.quantity, Decimal("100.000000"))
        self.assertEqual(row.amount, Decimal("5700.00"))

    def test_manual_pulp_ending_still_wins(self):
        """上传了浆粕期末（人工）时，仍以人工为准、不生成系统行。"""
        from settlement.services.costing import _persist_ending_by_batch

        manual = EndingInventoryBalance.objects.create(
            period=self.period, material_code=self.PULP_CODE,
            material_name=self.PULP_NAME, ending_date=date(2026, 7, 31),
            quantity=Decimal("999"), amount=Decimal("61938"),
            unit_price=Decimal("62"), is_system_generated=False,
        )
        _persist_ending_by_batch(self.period, self.summaries)
        self.assertFalse(self._system_ending().filter(material_code=self.PULP_CODE).exists())
        manual.refresh_from_db()
        self.assertEqual(manual.quantity, Decimal("999"))
        # 人工期末的价格会被 _fill_ending_prices 按月末加权价回填（数量不变）
        self.assertEqual(EndingInventoryBalance.objects.get(id=manual.id).quantity, Decimal("999"))

    def test_pulp_ending_is_carried_to_next_period(self):
        """浆粕期末能承接到下期期初（以前因为期末缺失而断档）。"""
        from settlement.services.costing import _persist_ending_by_batch
        from settlement.services.rollforward import _carry_inventory

        _persist_ending_by_batch(self.period, self.summaries)
        next_period = AccountingPeriod.objects.create(month=date(2026, 8, 1))
        _carry_inventory(next_period, self.period)
        opening = next_period.opening_inventory_balances.get(
            material_code=self.PULP_CODE
        )
        self.assertEqual(opening.quantity, Decimal("1300.000000"))
        self.assertEqual(opening.amount, Decimal("80600.00"))
        self.assertTrue(opening.is_system_generated)
        self.assertEqual(opening.source_period_id, self.period.id)
class EndingInventoryCategoryTests(SettlementFixtureMixin, TestCase):
    """期末存货行要带「分类名称」（用户 2026-09-19）。

    以前 `_persist_ending_by_batch` 生成期末行时分类写的是空串 —— 结果：
      · 期末存货导出/明细里「分类名称」列全空（看不出浆粕这类物料的分类）；
      · 下期承接会把空分类带到期初，若该物料下月没有流水，分类就彻底丢了（变「未分类」）。
    现在按 期初存货 → 本月流水 的顺序取第一个非空分类写进期末行。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="ending-category-user", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.batch = self.create_batch(
            self.period, ImportBatch.Kind.INVENTORY, "ending-category"
        )

    def _summaries(self, *codes):
        return {
            code: {
                "material_name": "分类原料" + code,
                "weighted_average_price": Decimal("10"),
            }
            for code in codes
        }

    def test_category_from_opening_inventory(self):
        """期初存货带分类 → 系统生成的期末行也带同一个分类。"""
        from settlement.services.costing import _persist_ending_by_batch

        OpeningInventoryBalance.objects.create(
            period=self.period, batch=self.batch, material_code="RM-CAT1",
            material_name="分类原料RM-CAT1", category_name="产成品",
            quantity=Decimal("10"), amount=Decimal("100"),
            unit_price=Decimal("10"), is_system_generated=False,
        )
        _persist_ending_by_batch(self.period, self._summaries("RM-CAT1"))
        row = EndingInventoryBalance.objects.get(
            period=self.period, material_code="RM-CAT1"
        )
        self.assertEqual(row.category_name, "产成品")
        self.assertEqual(row.quantity, Decimal("10.000000"))

    def test_category_falls_back_to_transaction(self):
        """期初没分类、本月流水有分类 → 用流水的分类。"""
        from settlement.services.costing import _persist_ending_by_batch

        material = Material.objects.create(
            code="RM-CAT2", name="分类原料RM-CAT2", uom="kg"
        )
        self.create_txn(
            self.period, self.batch, txn_date=date(2026, 7, 5),
            document_no="GR-CAT2", business_type="采购入库",
            direction=InventoryTransaction.Direction.IN, material=material,
            quantity=Decimal("20"),
        )
        txn = InventoryTransaction.objects.get(document_no="GR-CAT2")
        txn.category_name = "纤维"
        txn.save(update_fields=["category_name"])
        _persist_ending_by_batch(self.period, self._summaries("RM-CAT2"))
        row = EndingInventoryBalance.objects.get(
            period=self.period, material_code="RM-CAT2"
        )
        self.assertEqual(row.category_name, "纤维")

    def test_unclassified_stays_empty(self):
        """期初/流水都没有分类时：期末行分类留空（页面按「未分类」显示）。"""
        from settlement.services.costing import _persist_ending_by_batch
        from settlement import views as views_module

        OpeningInventoryBalance.objects.create(
            period=self.period, batch=self.batch, material_code="RM-CAT3",
            material_name="分类原料RM-CAT3", quantity=Decimal("3"),
            amount=Decimal("30"), unit_price=Decimal("10"),
            is_system_generated=False,
        )
        _persist_ending_by_batch(self.period, self._summaries("RM-CAT3"))
        row = EndingInventoryBalance.objects.get(
            period=self.period, material_code="RM-CAT3"
        )
        self.assertEqual(row.category_name, "")
        self.assertIsNone(
            views_module.inventory_category_map(self.period).get("RM-CAT3")
        )
class ContractAmountColumnTests(SettlementFixtureMixin, TestCase):
    """「到票与暂估」表新增列「合同金额」（用户 2026-09-19，加在期初暂估与本月到票中间）。

    口径：**归属本期**的采购合同不含税金额（优先文件的「原币未税总价（报表用）」，老数据退回 合同数量 × 不含税单价）；
      · 合同行填了「收货期间」→ 按收货期间归属；
      · 否则按该合同行所属期间的月份；
      · 按月「自动承接」的合同副本不重复计入（只由源头行算一次）。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="contract-amount-user", password="test-password"
        )
        self.client.force_login(self.user)
        self.june = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        self.july = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.contract_batch = self.create_batch(
            self.june, ImportBatch.Kind.PURCHASE_CONTRACT, "contract-amount"
        )
        self.material = Material.objects.create(
            code="RM-CAMT", name="合同金额原料", uom="kg"
        )

    def _contract(self, **overrides):
        values = dict(
            period=self.june,
            batch=self.contract_batch,
            contract_no="PC-CAMT",
            contract_date=date(2026, 6, 1),
            supplier_code="SUP-CAMT",
            supplier_name="合同金额供应商",
            material_code=self.material.code,
            material_name=self.material.name,
            quantity=Decimal("100"),
            remaining_quantity=Decimal("100"),
            unit_price_excl_tax=Decimal("10"),
        )
        values.update(overrides)
        return PurchaseContractLine.objects.create(**values)

    def _rows(self, period):
        resp = self.client.get(reverse("purchase_accruals", args=[period.id]))
        self.assertEqual(resp.status_code, 200)
        return {str(r["period"]): r for r in resp.context["flow_rows"]}, resp

    def test_falls_back_to_contract_period_when_no_receipt_period(self):
        """合同没填收货期间 → 合同金额落在合同所属期间（6 月）。"""
        self._contract()
        rows, _ = self._rows(self.july)
        self.assertEqual(rows[str(self.june)]["contract_amount"], Decimal("1000.00"))
        self.assertEqual(rows[str(self.july)]["contract_amount"], Decimal("0.00"))

    def test_follows_contract_receipt_period(self):
        """合同填了收货期间＝7 月 → 合同金额整个落在 7 月，6 月为 0。"""
        self._contract(receipt_period=date(2026, 7, 1), received_quantity=Decimal("100"))
        rows, resp = self._rows(self.july)
        self.assertEqual(rows[str(self.june)]["contract_amount"], Decimal("0.00"))
        self.assertEqual(rows[str(self.july)]["contract_amount"], Decimal("1000.00"))
        self.assertEqual(
            resp.context["flow_total"]["contract_amount"], Decimal("1000.00")
        )

    def test_carried_copy_is_not_counted_twice(self):
        """自动承接的合同副本不算第二次（否则 7 月会凭空多出 600）。"""
        original = self._contract()
        PurchaseContractLine.objects.create(
            period=self.july,
            carried_from=original,
            is_system_generated=True,
            contract_no=original.contract_no,
            contract_date=original.contract_date,
            supplier_code="SUP-CAMT",
            supplier_name="合同金额供应商",
            material_code=self.material.code,
            material_name=self.material.name,
            quantity=Decimal("60"),
            remaining_quantity=Decimal("60"),
            unit_price_excl_tax=Decimal("10"),
        )
        rows, resp = self._rows(self.july)
        self.assertEqual(rows[str(self.july)]["contract_amount"], Decimal("0.00"))
        self.assertEqual(
            resp.context["flow_total"]["contract_amount"], Decimal("1000.00")
        )

    def test_carry_forward_preserves_receipt_period(self):
        """承接要带上「收货期间／已签收数量」，否则承接后的月份匹配到副本 → 规则失效。"""
        from settlement.services.rollforward import _carry_contracts

        original = self._contract(
            receipt_period=date(2026, 7, 1), received_quantity=Decimal("100")
        )
        _carry_contracts(self.july, self.june)
        carried = PurchaseContractLine.objects.get(
            period=self.july, carried_from=original
        )
        self.assertTrue(carried.is_system_generated)
        self.assertEqual(carried.receipt_period, date(2026, 7, 1))
        self.assertEqual(carried.received_quantity, Decimal("100"))
        rows, _ = self._rows(self.july)
        self.assertEqual(rows[str(self.july)]["contract_amount"], Decimal("1000.00"))


class BasisBlockConsistencyTests(SettlementFixtureMixin, TestCase):
    """口径说明全站统一（用户 2026-09-21）。

    统一格式：左上角「口径说明」字样 + 一张四列表（字段 / 数据来源 / 计算依据 / 作用）。
    覆盖：采购发票 / 采购合同 / 月末暂估 / 暂估差异 / 暂估回冲 / 损益表 / 总成本表。
    """

    #: 有「口径说明」块的页面（url name → 该页面的口径块在哪）
    PAGES = (
        "purchase_matches",      # 各期间上传发票合计 下面
        "purchase_contracts",    # 采购合同汇总 下面
        "purchase_accruals",     # 到票与暂估 下面 + 暂估差异（2026-09-21 并入）三块
        # 注：生产成本页、月末在产品页、采购入库页的口径说明用「要点版」（.basis-notes，不是四列表）；
        #     损益表页 2026-09-28 也改成要点版，故这些页面不在本清单
    )

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="basis-block-user", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))

    def test_every_page_uses_the_same_basis_block(self):
        for name in self.PAGES:
            resp = self.client.get(reverse(name, args=[self.period.id]))
            self.assertEqual(resp.status_code, 200, name)
            html = resp.content.decode("utf-8")
            # ① 统一用 .basis-table 四列表，表头固定
            self.assertIn('class="basis-table"', html, name)
            self.assertIn("<th>数据来源</th><th>计算依据</th><th>作用</th>", html, name)
            self.assertIn("<th>字段</th>", html, name)
            # ② 左上角必须有「口径说明」字样（小标签，或该小节标题）
            self.assertTrue(
                'class="basis-title">口径说明<' in html or "<h2>口径说明</h2>" in html,
                name,
            )
            # ③ 旧样式名不再出现
            self.assertNotIn("accrual-basis-table", html, name)
            self.assertNotIn("note-list", html, name)

    def test_purchase_pages_show_the_small_label(self):
        """4 个采购页（+ 暂估回冲）的口径说明都带左上角小标签（不是小节标题）。"""
        for name in ("purchase_matches", "purchase_contracts", "purchase_accruals",
                     "purchase_receipts"):
            html = self.client.get(reverse(name, args=[self.period.id])).content.decode("utf-8")
            self.assertIn('class="basis-title">口径说明<', html, name)
            self.assertIn('class="basis-block"', html, name)

    def test_no_total_row_in_purchase_basis_tables(self):
        """用户 2026-09-21：采购发票 / 采购合同 的口径说明里不再列「合计」行。
        （注意：「发票不含税金额合计」这类格子里的词不算合计行。）"""
        for name in ("purchase_matches", "purchase_contracts"):
            html = self.client.get(reverse(name, args=[self.period.id])).content.decode("utf-8")
            block = html.split('class="basis-table"', 1)[1].split("</table>", 1)[0]
            self.assertNotIn("<strong>合计</strong>", block, name)
            self.assertEqual(block.count("<tr>"), 4, name)   # 1 行表头 + 3 行字段

    def test_reversal_detail_drops_the_period_pill(self):
        """用户 2026-09-21：「回冲明细（逐笔）」标题右侧的「2026年07月 · 本期」徽标删除，
        只剩「导出Excel」+（由 base.html 注入的）「打印本表」。"""
        html = self.client.get(
            reverse("purchase_accruals", args=[self.period.id])
        ).content.decode("utf-8")
        seg = html.split("<h2>回冲明细（逐笔）</h2>", 1)[1].split("</section>", 1)[0]
        self.assertNotIn("· 本期", seg)
        self.assertIn("导出Excel", seg)

    def test_basis_block_has_four_columns_and_rows(self):
        """每张口径表都是四列；每行四格都不为空。"""
        for name in self.PAGES:
            html = self.client.get(reverse(name, args=[self.period.id])).content.decode("utf-8")
            block = html.split('class="basis-table"', 1)[1].split("</table>", 1)[0]
            head = re.findall(r"<th[^>]*>([^<]+)</th>", block.split("</thead>")[0])
            self.assertEqual(head, ["字段", "数据来源", "计算依据", "作用"], name)
            rows = re.findall(r"<tr>(.*?)</tr>", block.split("</thead>", 1)[1], re.S)
            self.assertGreaterEqual(len(rows), 3, name)
            for row in rows:
                cells = re.findall(r"<td[^>]*>(.*?)</td>", row, re.S)
                self.assertEqual(len(cells), 4, "%s: %s" % (name, row[:80]))
                for cell in cells:
                    self.assertTrue(re.sub(r"<[^>]+>", "", cell).strip(), name)

    def test_profit_page_uses_concise_notes_not_a_table(self):
        """用户 2026-09-28：损益表页的口径说明改成简短要点，不再用四列表。"""
        html = self.client.get(
            reverse("profit_statement", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertIn('class="basis-title">口径说明<', html)
        self.assertIn('class="basis-notes"', html)
        self.assertNotIn("basis-table", html)
        notes = html.split('class="basis-notes"', 1)[1].split("</ul>", 1)[0]
        # 要点条数控制在 5 条以内（简洁优先）
        self.assertLessEqual(len(re.findall(r"<li>", notes)), 5)




class WorkOrderReportPageTests(SettlementFixtureMixin, TestCase):
    """「生产工单」页（用户 2026-09-21）：筛选栏整块删除。

    删掉的字段：工单编号 / 车间·部门 / 状态 + 「筛选」「清空」按钮。
    老链接/书签里带的这些参数仍然生效（视图照旧读），导出也会带过去。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="wo-report-user", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.batch = self.create_batch(
            self.period, ImportBatch.Kind.WORK_ORDER, "wo-report"
        )
        WorkOrder.objects.create(
            period=self.period, batch=self.batch, order_no="MO-RPT-1",
            order_type="普通", identifier="ID-1", plan_quantity=Decimal("100"),
            production_dept_code="tefang", status="执行中",
        )
        WorkOrder.objects.create(
            period=self.period, batch=self.batch, order_no="MO-RPT-2",
            order_type="普通", identifier="ID-2", plan_quantity=Decimal("50"),
            production_dept_code="RanZheng", status="已关闭",
        )
        self.url = reverse("work_order_report", args=[self.period.id])

    def test_filter_module_is_removed(self):
        html = self.client.get(self.url).content.decode("utf-8")
        # 三个筛选字段与筛选/清空按钮都不在了
        for name in ("order_no", "workshop", "status"):
            self.assertNotIn('name="%s"' % name, html, name)
        self.assertNotIn('class="filter-grid"', html)
        self.assertNotIn(">清空<", html)
        self.assertNotIn(">筛选<", html)
        # 页面主体仍在：表格列＝上传文件的 25 列（列名与顺序一致）
        self.assertIn("<h1>生产工单</h1>", html)
        # 用户 2026-09-21：导出Excel 从页头挪到表格标题右侧（与「打印本表」同一排）
        page_head = html.split("</section>", 1)[0]
        self.assertNotIn("导出Excel", page_head)
        seg_head = html.split("<h2>生产工单明细</h2>", 1)[1].split("</table>", 1)[0]
        self.assertIn("导出Excel", seg_head)
        self.assertIn('class="print-here"', seg_head)
        self.assertLess(
            seg_head.index("导出Excel"), seg_head.index('class="print-here"')
        )
        heads = [re.sub(r"<[^>]+>", "", h).strip() for h in re.findall(
            r"<th[^>]*>(.*?)</th>", html, re.S)]
        self.assertEqual(
            heads,
            ["工单编号", "工单类型", "标识码", "计划数量", "计划开始时间",
             "计划完工时间", "优先级", "产线设备", "班次", "生产部门编号",
             "生产主管编号", "计划部门编号", "计划员编号", "备注", "指定用料",
             "外部单据编号", "状态", "是否使用物料清单", "是否使用工艺路线",
             "工序计划类型", "启用SOP", "生产模式", "产线编号", "订单号", "物料颜色"],
        )
        self.assertEqual(len(list(self.client.get(self.url).context["page"])), 2)
        # 每行 25 格
        seg = html.split("<tbody>", 1)[1].split("</tbody>", 1)[0]
        for row in re.findall(r"<tr>(.*?)</tr>", seg, re.S):
            self.assertEqual(len(re.findall(r"<td[^>]*>", row)), 25, row[:60])
        # 上传文件里的写法：计划数量不带小数尾巴、布尔列写 是/否、
        # 生产部门编号显示原始编号（不再翻成中文车间名）
        self.assertIn(">100<", html)          # 计划数量 100（不是 100.000）
        self.assertNotIn("100.000", html)
        self.assertIn(">tefang<", html)
        self.assertIn(">否<", html)           # use_bom 默认 False → 否

    def test_legacy_query_params_still_work(self):
        """老链接里带的参数仍然生效（只是页面上看不到筛选栏了）。"""
        by_no = self.client.get(self.url, {"order_no": "MO-RPT-1"}).context["page"]
        self.assertEqual([o.order_no for o in by_no], ["MO-RPT-1"])
        by_status = self.client.get(self.url, {"status": "已关闭"}).context["page"]
        self.assertEqual([o.order_no for o in by_status], ["MO-RPT-2"])
        by_ws = self.client.get(self.url, {"workshop": "tefang"}).context["page"]
        self.assertEqual([o.order_no for o in by_ws], ["MO-RPT-1"])
        # 导出也照旧带上参数（全量导出＝页面所见），且与页面一样是 25 列
        resp = self.client.get(
            reverse("export_listing", args=[self.period.id, "work_order"]),
            {"order_no": "MO-RPT-2"},
        )
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(BytesIO(resp.content), read_only=True)
        try:
            ws = wb[wb.sheetnames[0]]
            head = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
            rows = list(ws.iter_rows(min_row=2, values_only=True))
        finally:
            wb.close()
        self.assertEqual(len(head), 25)
        self.assertEqual(head[:4], ["工单编号", "工单类型", "标识码", "计划数量"])
        self.assertEqual(head[-1], "物料颜色")
        self.assertEqual(len(rows), 1)                      # 按工单号筛出 1 行
        self.assertEqual(rows[0][0], "MO-RPT-2")
        self.assertEqual(rows[0][3], Decimal("50"))         # 计划数量仍是数字


class EndingWipPageTests(SettlementFixtureMixin, TestCase):
    """「期末在产品」页（用户 2026-09-21 新增，成本核算组）。

    · 数据来源＝工单结算结果（WorkOrderCostResult）的 wip_* 三个科目；
    · 只列**还有在产品**的工单（在产品数量或在产品成本任一不为 0）；
    · 给出「各车间期末在产品」汇总 + 按工单明细；可筛选车间 / 产品编码或名称；可导出。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="wip-page-user", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        # 纺纱车间两张有在产品的工单
        self.wo_a = WorkOrder.objects.create(
            period=self.period, order_no="WIP-A", status="执行中",
            production_dept_code="tefang", plan_quantity=Decimal("100"),
        )
        WorkOrderCostResult.objects.create(
            period=self.period, work_order=self.wo_a,
            product_code="P1", product_name="粗纱", plan_quantity=Decimal("100"),
            finished_quantity=Decimal("80"), wip_quantity=Decimal("20"),
            wip_material=Decimal("1000"), wip_wages=Decimal("200"),
            wip_overhead=Decimal("300"), wip_cost=Decimal("1500"),
            finished_cost=Decimal("3500"), total_cost=Decimal("5000"),
        )
        self.wo_c = WorkOrder.objects.create(
            period=self.period, order_no="WIP-C", status="已下发",
            production_dept_code="tefang", plan_quantity=Decimal("50"),
        )
        WorkOrderCostResult.objects.create(
            period=self.period, work_order=self.wo_c,
            product_code="P2", product_name="细纱", plan_quantity=Decimal("50"),
            finished_quantity=Decimal("40"), wip_quantity=Decimal("10"),
            wip_material=Decimal("500"), wip_wages=Decimal("100"),
            wip_overhead=Decimal("0"), wip_cost=Decimal("600"),
            finished_cost=Decimal("900"), total_cost=Decimal("1500"),
        )
        # 机织车间：已全部完工（在产品为 0）→ 不应出现在本页
        self.wo_b = WorkOrder.objects.create(
            period=self.period, order_no="WIP-B", status="已完成",
            production_dept_code="JiZhi", plan_quantity=Decimal("30"),
        )
        WorkOrderCostResult.objects.create(
            period=self.period, work_order=self.wo_b,
            product_code="P3", product_name="坯布", plan_quantity=Decimal("30"),
            finished_quantity=Decimal("30"), wip_quantity=Decimal("0"),
            wip_material=Decimal("0"), wip_wages=Decimal("0"),
            wip_overhead=Decimal("0"), wip_cost=Decimal("0"),
            finished_cost=Decimal("800"), total_cost=Decimal("800"),
        )
        self.url = reverse("ending_wip", args=[self.period.id])

    def test_lists_only_orders_with_wip(self):
        resp = self.client.get(self.url)
        self.assertEqual(resp.status_code, 200)
        rows = list(resp.context["page"])
        # 只有两张有在产品的工单；已完工那张不在
        self.assertEqual([row["order_no"] for row in rows], ["WIP-A", "WIP-C"])
        first = rows[0]
        self.assertEqual(first["workshop"], "纺纱车间")
        self.assertEqual(first["product_code"], "P1")
        self.assertEqual(first["wip_quantity"], Decimal("20"))
        self.assertEqual(first["wip_material"], Decimal("1000.00"))
        self.assertEqual(first["wip_wages"], Decimal("200.00"))
        self.assertEqual(first["wip_overhead"], Decimal("300.00"))
        self.assertEqual(first["wip_cost"], Decimal("1500.00"))
        # 明细合计 = 各工单在产品成本之和（且只算有在产品的）
        self.assertEqual(
            resp.context["total"]["amount"],
            sum((r.wip_cost for r in WorkOrderCostResult.objects.filter(
                period=self.period)), Decimal("0")),
        )

    def test_workshop_summary_and_total(self):
        resp = self.client.get(self.url)
        html = resp.content.decode("utf-8")
        heads = re.findall(
            r"<th[^>]*>([^<]+)</th>",
            html.split('class="wip-summary-table"', 1)[1].split("</thead>")[0],
        )
        self.assertEqual(
            heads,
            ["车间", "在产品工单数", "在产品数量", "直接材料", "人工费",
             "制造费用", "在产品成本"],
        )
        rows = resp.context["workshop_rows"]
        # 机织车间在产品为 0 → 不出行；只剩纺纱一行
        self.assertEqual(len(rows), 1)
        tefang = rows[0]
        self.assertEqual(tefang["workshop"], "纺纱车间")
        self.assertEqual(tefang["workshop_code"], "tefang")
        self.assertEqual(tefang["order_count"], 2)
        self.assertEqual(tefang["wip_quantity"], Decimal("30"))
        self.assertEqual(tefang["wip_material"], Decimal("1500.00"))
        self.assertEqual(tefang["wip_wages"], Decimal("300.00"))
        self.assertEqual(tefang["wip_overhead"], Decimal("300.00"))
        self.assertEqual(tefang["amount"], Decimal("2100.00"))
        total = resp.context["total"]
        self.assertEqual(total["order_count"], 2)
        self.assertEqual(total["wip_quantity"], Decimal("30"))
        self.assertEqual(total["wip_material"], Decimal("1500.00"))
        self.assertEqual(total["wip_wages"], Decimal("300.00"))
        self.assertEqual(total["wip_overhead"], Decimal("300.00"))
        self.assertEqual(total["amount"], Decimal("2100.00"))
        # 汇总 = 明细三科目相加（不重不漏）
        self.assertEqual(
            total["amount"],
            total["wip_material"] + total["wip_wages"] + total["wip_overhead"],
        )
        self.assertIn("合计", html.split('class="wip-summary-table"', 1)[1])

    def test_page_columns_filters_and_notes(self):
        resp = self.client.get(self.url)
        html = resp.content.decode("utf-8")
        self.assertIn("<h1>期末在产品</h1>", html)
        for heading in ("各车间期末在产品", "期末在产品明细（按工单）"):
            self.assertIn("<h2>%s</h2>" % heading, html, heading)
        # 明细表头 12 列、每行 12 格
        heads = [re.sub(r"<[^>]+>", "", h).strip() for h in re.findall(
            r"<th[^>]*>(.*?)</th>",
            html.split("<h2>期末在产品明细（按工单）</h2>", 1)[1].split("</thead>")[0],
            re.S,
        )]
        # 用户 2026-09-21：「在产品成本」列已从明细表删除 → 12 列
        self.assertEqual(len(heads), 12)
        self.assertEqual(
            heads[:5], ["车间", "起始期间", "工单编号", "工单状态", "产品编码"]
        )
        self.assertIn("在产品数量", heads)
        self.assertNotIn("在产品成本", heads)
        self.assertEqual(heads[-1], "制造费用")
        seg = html.split("<h2>期末在产品明细（按工单）</h2>", 1)[1].split("</table>", 1)[0]
        for row in re.findall(r"<tr>(.*?)</tr>", seg.split("</thead>", 1)[1], re.S):
            self.assertEqual(len(re.findall(r"<td[^>]*>", row)), 12, row[:80])
        # 口径说明：要点版（不用表格）+ 左上角标签
        self.assertIn('class="basis-title">口径说明<', html)
        self.assertIn('class="basis-notes"', html)
        self.assertNotIn('class="basis-table"', html)
        # 口径说明要点（用户 2026-09-21 要求简洁）：压缩成 3 条
        self.assertEqual(
            len(re.findall(r"<li>", html.split('class="basis-notes"', 1)[1].split("</ul>", 1)[0])),
            3,
        )
        # 筛选：车间下拉（只有有在产品的车间）+ 产品编码输入
        self.assertIn('name="workshop"', html)
        self.assertIn('name="material"', html)
        # 用户 2026-09-21：车间下拉选完就自动提交（否则选了没反应）；
        # 但**不要**做任何滚动跳转 —— 保持页面不动（base.html 会还原滚动位置）。
        self.assertIn('name="workshop" onchange="this.form.submit()"', html)
        self.assertNotIn("scrollIntoView", html)
        self.assertNotIn("#wip-detail", html)
        summary = html.split('class="wip-summary-table"', 1)[1].split("</table>", 1)[0]
        # 车间名可点（跳到该车间的筛选结果），但不带锚点（避免页面跳动）
        self.assertIn("?workshop=tefang", summary)

        # 筛选生效：车间=机织（在产品为 0）→ 空；产品=P1 → 1 行
        self.assertEqual(list(self.client.get(
            self.url, {"workshop": "JiZhi"}).context["page"]), [])
        self.assertEqual(
            [r["order_no"] for r in self.client.get(
                self.url, {"material": "P1"}).context["page"]],
            ["WIP-A"],
        )
        # 车间筛纺纱 → 2 行；汇总随筛选变化
        tefang = self.client.get(self.url, {"workshop": "tefang"})
        self.assertEqual(len(list(tefang.context["page"])), 2)
        self.assertEqual(tefang.context["total"]["amount"], Decimal("2100.00"))

    def test_start_period_column(self):
        """用户 2026-09-21：明细加「起始期间」＝这张工单什么时候开始有的在产品。

        · 期初台账里就有的工单 → 「期初结转」（早于本期）；
        · 本期才出现的工单 → 本期（YYYY-MM）。
        （汇总的「期初在产品成本」列按用户后续要求已删除。）
        """
        # WIP-A 在本期期初台账里已有在产品（材料 800 + 人工 200）
        OpeningWip.objects.create(
            period=self.period, order_no="WIP-A", quantity=Decimal("5"),
            material=Decimal("800"), wages=Decimal("200"),
        )
        resp = self.client.get(self.url)
        rows = {r["workshop_code"]: r for r in resp.context["workshop_rows"]}
        # 汇总只按期末在产品出车间（没有期初列了）
        self.assertNotIn("opening_amount", rows["tefang"])
        self.assertEqual(rows["tefang"]["amount"], Decimal("2100.00"))

        detail = {r["order_no"]: r for r in resp.context["page"]}
        self.assertEqual(detail["WIP-A"]["start_period"], "期初结转")
        self.assertEqual(detail["WIP-C"]["start_period"], "2026-07")
        html = resp.content.decode("utf-8")
        self.assertIn("期初结转", html)
        # 口径说明讲清「起始期间」（要点版）
        notes = html.split('class="basis-notes"', 1)[1].split("</ul>", 1)[0]
        self.assertIn("起始期间", notes)

        # 导出两张表都带上新列
        resp2 = self.client.get(
            reverse("export_listing", args=[self.period.id, "ending_wip"])
        )
        wb = load_workbook(BytesIO(resp2.content), read_only=True)
        try:
            head = [c.value for c in next(
                wb["期末在产品"].iter_rows(min_row=1, max_row=1))]
            self.assertEqual(head[:3], ["车间", "起始期间", "工单编号"])
            head2 = [c.value for c in next(
                wb["按车间汇总"].iter_rows(min_row=1, max_row=1))]
            self.assertNotIn("期初在产品成本", head2)
        finally:
            wb.close()

    def test_summary_stays_full_when_detail_is_filtered(self):
        """明细的车间筛选不带走汇总表（汇总始终全量），金额徽标显示筛选后合计。"""
        full = self.client.get(self.url)
        filtered = self.client.get(self.url, {"workshop": "JiZhi"})
        self.assertEqual(
            [(r["workshop_code"], r["amount"]) for r in filtered.context["workshop_rows"]],
            [(r["workshop_code"], r["amount"]) for r in full.context["workshop_rows"]],
        )
        self.assertEqual(filtered.context["total"], full.context["total"])
        # 明细被筛空（fixture 里机织那张已完工）
        self.assertEqual(list(filtered.context["page"]), [])
        self.assertEqual(filtered.context["filtered_total"]["amount"], Decimal("0.00"))

    def test_nav_has_the_entry_and_export_has_two_sheets(self):
        html = self.client.get(
            reverse("period_detail", args=[self.period.id])
        ).content.decode("utf-8")
        nav = html.split('data-nav="production"', 1)[1].split("</div>", 2)[0]
        self.assertIn("期末在产品", nav)
        self.assertIn(reverse("ending_wip", args=[self.period.id]), nav)
        self.assertLess(nav.index("生产成本"), nav.index("期末在产品"))

        resp = self.client.get(
            reverse("export_listing", args=[self.period.id, "ending_wip"])
        )
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(BytesIO(resp.content), read_only=True)
        try:
            self.assertEqual(wb.sheetnames, ["期末在产品", "按车间汇总"])
            ws = wb["期末在产品"]
            head = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
            self.assertEqual(
                head[:5], ["车间", "起始期间", "工单编号", "工单状态", "产品编码"]
            )
            self.assertEqual(len(head), 12)
            self.assertNotIn("在产品成本", head)      # 已按用户要求从明细删除
            rows = list(ws.iter_rows(min_row=2, values_only=True))
            self.assertEqual(rows[0][2], "WIP-A")       # 第 3 列＝工单编号（前面是起始期间）
            self.assertEqual(rows[-1][0], "合计")          # 末行合计
            self.assertEqual(rows[-1][3], None)
            ws2 = wb["按车间汇总"]
            head2 = [c.value for c in next(ws2.iter_rows(min_row=1, max_row=1))]
            self.assertEqual(
                head2, ["车间", "在产品工单数", "在产品数量", "直接材料", "人工费",
                        "制造费用", "在产品成本"],
            )
            rows2 = list(ws2.iter_rows(min_row=2, values_only=True))
            self.assertEqual(rows2[0][0], "纺纱车间")
            self.assertEqual(rows2[-1][0], "合计")
        finally:
            wb.close()

class PurchaseReceiptPageTests(SettlementFixtureMixin, TestCase):
    """「采购入库」页（用户 2026-09-21 新增）：本期间录入的采购入库流水。

    · 数据＝业务类型「采购入库」且**流水所属期间＝本期**的流水（与月末暂估按归属月份不同）；
    · 每行给出数量、暂估单价/金额、累计到票、尚未到票、状态；
    · 上面一张「各期间采购入库合计」；可导出《采购入库》。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="receipt-page-user", password="test-password"
        )
        self.client.force_login(self.user)
        self.june = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        self.july = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.inventory_batch = self.create_batch(
            self.june, ImportBatch.Kind.INVENTORY, "receipt-inv"
        )
        self.july_inventory_batch = self.create_batch(
            self.july, ImportBatch.Kind.INVENTORY, "receipt-inv-july"
        )
        self.contract_batch = self.create_batch(
            self.june, ImportBatch.Kind.PURCHASE_CONTRACT, "receipt-contract"
        )
        self.invoice_batch = self.create_batch(
            self.july, ImportBatch.Kind.PURCHASE_INVOICE, "receipt-invoice"
        )
        self.material = Material.objects.create(
            code="RM-RCPT", name="入库测试原料", uom="kg"
        )
        PurchaseContractLine.objects.create(
            period=self.june, batch=self.contract_batch, contract_no="PC-RCPT",
            contract_date=date(2026, 6, 1), supplier_code="SUP-RCPT",
            supplier_name="入库供应商", material_code=self.material.code,
            material_name=self.material.name, quantity=Decimal("200"),
            remaining_quantity=Decimal("200"), unit_price_excl_tax=Decimal("10"),
        )
        # 6 月到货 100@10（合同暂估）
        self.create_txn(
            self.june, self.inventory_batch, txn_date=date(2026, 6, 5),
            document_no="GR-RCPT-06", business_type="采购入库",
            direction=InventoryTransaction.Direction.IN, material=self.material,
            quantity=Decimal("100"), supplier_code="SUP-RCPT",
            supplier_name="入库供应商", contract_no="PC-RCPT",
        )
        # 7 月到货 50@10 + 一张 7 月发票核销 30@10 → 累计到票 30、尚未到票 20
        # （7 月用**另一个物料/合同**，保证发票只能核销到 7 月这单，不会挂到 6 月那单上）
        self.material2 = Material.objects.create(
            code="RM-RCPT2", name="入库测试原料二", uom="kg"
        )
        PurchaseContractLine.objects.create(
            period=self.july, batch=self.contract_batch, contract_no="PC-RCPT2",
            contract_date=date(2026, 7, 1), supplier_code="SUP-RCPT",
            supplier_name="入库供应商", material_code=self.material2.code,
            material_name=self.material2.name, quantity=Decimal("50"),
            remaining_quantity=Decimal("50"), unit_price_excl_tax=Decimal("10"),
        )
        self.create_txn(
            self.july, self.july_inventory_batch, txn_date=date(2026, 7, 6),
            document_no="GR-RCPT-07", business_type="采购入库",
            direction=InventoryTransaction.Direction.IN, material=self.material2,
            quantity=Decimal("50"), supplier_code="SUP-RCPT",
            supplier_name="入库供应商", contract_no="PC-RCPT2",
        )
        PurchaseInvoiceLine.objects.create(
            period=self.july, batch=self.invoice_batch, invoice_no="PI-RCPT-07",
            invoice_date=date(2026, 7, 25), invoice_period=date(2026, 7, 1),
            contract_no="PC-RCPT2", supplier_code="SUP-RCPT",
            supplier_name="入库供应商", material_code=self.material2.code,
            material_name=self.material2.name, quantity=Decimal("30"),
            remaining_quantity=Decimal("30"), unit_price_excl_tax=Decimal("10"),
            amount_excl_tax=Decimal("300"),
        )
        auto_match_all(self.july)
        self.url = reverse("purchase_receipts", args=[self.july.id])

    def test_page_lists_all_receipts_with_columns(self):
        """默认列「全部期间」的采购入库流水（用户 2026-09-21 去掉了「本期」筛选）。"""
        resp = self.client.get(self.url)
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        self.assertIn("<h1>采购入库</h1>", html)
        for heading in ("采购汇总", "采购入库明细"):
            self.assertIn("<h2>%s</h2>" % heading, html, heading)
        # 旧标题不应再出现
        self.assertNotIn("各期间采购入库合计", html)
        self.assertNotIn("本期采购入库明细", html)
        rows = list(resp.context["page"])
        # 默认全部期间：6 月、7 月两单都在（按入库日期排序）
        self.assertEqual(
            [row["document_no"] for row in rows], ["GR-RCPT-06", "GR-RCPT-07"]
        )
        row = rows[1]
        self.assertEqual(row["quantity"], Decimal("50"))
        self.assertEqual(row["provisional_unit_price"], Decimal("10.00"))
        self.assertEqual(row["provisional_amount"], Decimal("500.00"))
        self.assertEqual(row["invoiced_quantity"], Decimal("30"))
        self.assertEqual(row["outstanding_quantity"], Decimal("20"))
        self.assertEqual(row["status"], "部分到票")
        self.assertEqual(row["latest_invoice_period"], date(2026, 7, 1))
        # 「入库期间」取入库日期的年月（用户 2026-09-21），且不留空
        self.assertEqual(row["receipt_month_label"], "2026-07")
        self.assertEqual(row["receipt_month"], date(2026, 7, 1))

    def test_period_summary_and_columns(self):
        resp = self.client.get(self.url)
        html = resp.content.decode("utf-8")
        # 「采购汇总」按「入库年月 × 物料分类」出行（用户 2026-09-21），与明细同口径；
        # 列＝入库期间 / 物料分类 / 入库数量 / 暂估金额
        summary_heads = re.findall(
            r"<th[^>]*>([^<]+)</th>",
            html.split('class="receipt-summary-table"', 1)[1].split("</thead>")[0],
        )
        self.assertEqual(
            summary_heads, ["入库期间", "物料分类", "入库数量", "暂估金额"]
        )
        summary = {
            (row["month_label"], row["category"]): row
            for row in resp.context["summary_rows"]
        }
        # 测试数据没给物料分类 → 记「未分类」
        july = summary[("2026-07", "未分类")]
        self.assertEqual(july["quantity"], Decimal("50"))
        self.assertEqual(july["amount"], Decimal("500.00"))
        self.assertTrue(july["is_current"])
        self.assertEqual(july["period"], self.july)
        self.assertEqual(summary[("2026-06", "未分类")]["quantity"], Decimal("100"))
        # 已删除的字段不再出现
        for gone in ("document_count", "invoiced_quantity", "outstanding_quantity"):
            self.assertNotIn(gone, july)
        # 明细表：18 列（2026-09-21 新增「入库期间」与「业务类型」列），每行列数一致
        detail = html.split("<h2>采购入库明细</h2>", 1)[1].split("</table>", 1)[0]
        heads = re.findall(r"<th[^>]*>([^<]+)</th>", detail.split("</thead>")[0])
        self.assertEqual(len(heads), 18)
        self.assertEqual(
            heads[:5], ["入库期间", "入库日期", "入库单号", "业务类型", "供应商"]
        )
        self.assertIn("物料编码", heads)
        self.assertIn("物料名称", heads)
        for row in re.findall(r"<tr>(.*?)</tr>", detail.split("</thead>", 1)[1], re.S):
            self.assertEqual(len(re.findall(r"<td[^>]*>", row)), 18, row[:80])
        # 汇总表结构 + 口径说明（要点版，不用表格）
        self.assertIn('class="receipt-summary-table"', html)
        self.assertIn('class="basis-title">口径说明<', html)
        self.assertIn('class="basis-notes"', html)
        self.assertNotIn('class="basis-table"', html)

    def test_summary_splits_by_material_category(self):
        """用户 2026-09-21：采购汇总加「物料分类」——同一个入库期间按分类拆成多行。"""
        category_map = {
            self.material.code: "原料",       # 6 月那单
            self.material2.code: "纱线",      # 7 月那单
        }
        for code, name in category_map.items():
            InventoryTransaction.objects.filter(
                period__in=(self.june, self.july), material_code=code
            ).update(category_name=name)

        resp = self.client.get(self.url)
        rows = {
            (row["month_label"], row["category"]): row
            for row in resp.context["summary_rows"]
        }
        self.assertEqual(
            sorted(rows), [("2026-06", "原料"), ("2026-07", "纱线")]
        )
        self.assertEqual(rows[("2026-06", "原料")]["quantity"], Decimal("100"))
        self.assertEqual(rows[("2026-06", "原料")]["amount"], Decimal("1000.00"))
        self.assertEqual(rows[("2026-07", "纱线")]["quantity"], Decimal("50"))
        self.assertEqual(rows[("2026-07", "纱线")]["amount"], Decimal("500.00"))
        # 表头与页面列数一致（4 列），每行 4 格
        html = resp.content.decode("utf-8")
        table = html.split('class="receipt-summary-table"', 1)[1].split("</table>", 1)[0]
        heads = re.findall(r"<th[^>]*>([^<]+)</th>", table.split("</thead>")[0])
        self.assertEqual(heads, ["入库期间", "物料分类", "入库数量", "暂估金额"])
        for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", table.split("</thead>", 1)[1], re.S):
            self.assertEqual(len(re.findall(r"<td[^>]*>", tr)), 4, tr[:80])
        # 汇总合计仍等于明细合计（拆分不改变总数）
        self.assertEqual(
            sum(row["quantity"] for row in resp.context["summary_rows"]),
            sum(row["quantity"] for row in resp.context["page"]),
        )

    def test_detail_filters_by_period_supplier_material(self):
        """用户 2026-09-21：明细支持 入库期间 / 供应商 / 物料编码 筛选。

        · 入库期间：默认「全部期间」；可切到具体年月；
          「本期」那个选项已按要求删除；
        · 供应商、物料编码按「编码或名称」模糊匹配；
        · 只影响明细与导出，上面的「采购汇总」也跟随同一口径（同一批数据）。
        """
        base = self.client.get(self.url)
        html = base.content.decode("utf-8")

        # 默认＝全部期间；选项里只有 全部期间 + 各入库年月，没有「本期」
        self.assertEqual(base.context["receipt_filter"]["period"], "all")
        options = re.findall(r'<option value="([^"]*)"', html)
        self.assertIn("all", options)
        self.assertNotIn("current", options)
        self.assertNotIn("本期", html.split('name="receipt_period"', 1)[1].split("</select>")[0])
        # 「入库类型」筛选字段已删除
        self.assertNotIn('name="type"', html)

        # 默认（全部期间）：6 月那单与 7 月那单都在
        self.assertEqual(
            sorted(row["document_no"] for row in base.context["page"]),
            ["GR-RCPT-06", "GR-RCPT-07"],
        )

        # 入库期间＝2026-07：只剩 7 月那单
        july_rows = self.client.get(self.url, {"receipt_period": "2026-07"}).context["page"]
        self.assertEqual([row["document_no"] for row in july_rows], ["GR-RCPT-07"])
        # 入库期间＝2026-06：只剩 6 月那单（到票进度仍截至本期）
        june_rows = self.client.get(self.url, {"receipt_period": "2026-06"}).context["page"]
        self.assertEqual([row["document_no"] for row in june_rows], ["GR-RCPT-06"])

        # 供应商 / 物料编码
        by_supplier = self.client.get(self.url, {"supplier": "入库供应商"}).context["page"]
        self.assertEqual(len(list(by_supplier)), 2)
        self.assertEqual(
            list(self.client.get(
                self.url, {"supplier": "查无此供应商"}
            ).context["page"]),
            [],
        )
        by_material = self.client.get(self.url, {"material": "RM-RCPT2"}).context["page"]
        self.assertEqual([row["document_no"] for row in by_material], ["GR-RCPT-07"])
        self.assertEqual(
            list(self.client.get(self.url, {"material": "ZZZ"}).context["page"]), []
        )

        # 「采购汇总」是全量口径：不随明细筛选变化（筛不到时分项为空）
        full_summary = [
            (r["month_label"], r["category"], r["quantity"])
            for r in base.context["summary_rows"]
        ]
        self.assertEqual(
            full_summary,
            [("2026-06", "未分类", Decimal("100")),
             ("2026-07", "未分类", Decimal("50"))],
        )
        filtered = self.client.get(self.url, {"supplier": "查无此供应商"})
        self.assertEqual(list(filtered.context["page"]), [])
        self.assertEqual(
            [(r["month_label"], r["category"], r["quantity"])
             for r in filtered.context["summary_rows"]],
            full_summary,
        )
        # 导出跟随筛选：按物料筛后只剩 1 行
        resp = self.client.get(
            reverse("export_listing", args=[self.july.id, "purchase_receipt"]),
            {"material": "RM-RCPT2"},
        )
        wb = load_workbook(BytesIO(resp.content), read_only=True)
        try:
            ws = wb["采购入库"] if "采购入库" in wb.sheetnames else wb[wb.sheetnames[0]]
            rows = list(ws.iter_rows(min_row=2, values_only=True))
        finally:
            wb.close()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][2], "GR-RCPT-07")

    def test_including_returns_reconciles_with_inventory_rollforward(self):
        """用户追问：为什么采购入库的原料与库存结转的原料不一样？

        原因：「库存结转」把**所有入库类型**都计进入库数量/金额，其中包含
        「退货入库」（客户退货回仓）；本页默认只看「采购入库」。
        把「入库类型」切到「含退货入库」，同一物料分类的入库数量/金额就与库存结转完全一致。
        """
        # 造一条客户退货回仓（退货入库）：原料类物料
        self.create_txn(
            self.july, self.july_inventory_batch, txn_date=date(2026, 7, 9),
            document_no="RT-RCPT-01", business_type="退货入库",
            direction=InventoryTransaction.Direction.IN, material=self.material,
            quantity=Decimal("20"),
        )

        # 默认（只采购入库）：不含退货行；数量＝6 月 100 ＋ 7 月 50
        base = self.client.get(self.url, {"receipt_period": "all"})
        base_rows = list(base.context["page"])
        self.assertEqual(
            [r for r in base_rows if r["business_type"] == "退货入库"], []
        )
        self.assertEqual(
            sum((r["quantity"] for r in base_rows), Decimal("0")), Decimal("150")
        )
        self.assertEqual({r["business_type"] for r in base_rows}, {"采购入库"})
        self.assertEqual(
            [r["document_no"] for r in base_rows], ["GR-RCPT-06", "GR-RCPT-07"]
        )

        # 含退货入库：带上退货行，业务类型列区分得出来
        with_returns = self.client.get(
            self.url, {"receipt_period": "all", "type": "with_return"}
        )
        rows = list(with_returns.context["page"])
        self.assertEqual(
            [r["document_no"] for r in rows if r["business_type"] == "退货入库"],
            ["RT-RCPT-01"],
        )
        self.assertEqual(sum((r["quantity"] for r in rows), Decimal("0")), Decimal("170"))
        # 汇总表也跟着口径变：7 月从 50 变 70
        labels = {r["month_label"]: r for r in with_returns.context["summary_rows"]}
        self.assertEqual(labels["2026-07"]["quantity"], Decimal("70"))
        self.assertEqual(labels["2026-06"]["quantity"], Decimal("100"))
        # 采购汇总只保留三列（该口径下 7 月同样多 20）
        self.assertEqual(labels["2026-07"]["amount"], Decimal("700.00"))

        # 页面上的业务类型列：退货行带「退货入库」徽标
        html = with_returns.content.decode("utf-8")
        seg = html.split("<h2>采购入库明细</h2>", 1)[1].split("</table>", 1)[0]
        self.assertIn("退货入库", seg)
        # 页面上已无「入库类型」筛选（用户 2026-09-21 要求删除），
        # 需要与库存结转对账时仍可用 URL 参数 ?type=with_return
        self.assertNotIn('name="type"', html)

        # 导出也带该行与业务类型列
        resp = self.client.get(
            reverse("export_listing", args=[self.july.id, "purchase_receipt"]),
            {"receipt_period": "all", "type": "with_return"},
        )
        wb = load_workbook(BytesIO(resp.content), read_only=True)
        try:
            ws = wb["采购入库"] if "采购入库" in wb.sheetnames else wb[wb.sheetnames[0]]
            rows2 = list(ws.iter_rows(min_row=2, values_only=True))
        finally:
            wb.close()
        self.assertEqual(len(rows2), 3)          # 6 月 1 单 + 7 月 1 单 + 退货 1 单
        self.assertIn("退货入库", [r[3] for r in rows2])

    def test_receipt_month_comes_from_txn_date_not_period(self):
        """用户 2026-09-21：入库期间＝入库日期的年月（不是上传所属期间），且不留空。

        构造一张**补录进 7 月、但单据日期是 2026-06-30** 的采购入库：
        · 明细里它的「入库期间」必须是 2026-06（所属期间仍是 7 月）；
        · 入库期间筛 2026-06 能找到它；筛 2026-07 则不含它；
        · 页面与导出的「入库期间」列都是 YYYY-MM 文本，没有空值。
        """
        self.create_txn(
            self.july, self.july_inventory_batch, txn_date=date(2026, 6, 30),
            document_no="GR-RCPT-LATE", business_type="采购入库",
            direction=InventoryTransaction.Direction.IN, material=self.material,
            quantity=Decimal("10"), supplier_code="SUP-RCPT",
            supplier_name="入库供应商", contract_no="PC-RCPT",
        )
        resp = self.client.get(self.url)          # 默认＝全部期间
        rows = {row["document_no"]: row for row in resp.context["page"]}
        self.assertEqual(rows["GR-RCPT-LATE"]["receipt_month_label"], "2026-06")
        self.assertEqual(rows["GR-RCPT-LATE"]["receipt_month"], date(2026, 6, 1))
        # 所属期间仍是 7 月（说明差异确实来自口径，而不是数据没进库）
        self.assertEqual(rows["GR-RCPT-LATE"]["receipt_period"], self.july)

        # 按入库期间 2026-06 筛：能找到这张补录单
        june = self.client.get(self.url, {"receipt_period": "2026-06"})
        self.assertIn("GR-RCPT-LATE", {r["document_no"] for r in june.context["page"]})
        # 按 2026-07 筛：不含它
        july = self.client.get(self.url, {"receipt_period": "2026-07"})
        self.assertNotIn("GR-RCPT-LATE", {r["document_no"] for r in july.context["page"]})

        # 页面上的「入库期间」列没有空值，都是 YYYY-MM
        html = resp.content.decode("utf-8")
        detail = html.split("<h2>采购入库明细</h2>", 1)[1].split("</table>", 1)[0]
        first_cells = re.findall(r"<tr>\s*<td><strong>(.*?)</strong></td>", detail, re.S)
        self.assertTrue(first_cells)
        for cell in first_cells:
            self.assertRegex(cell.strip(), r"^\d{4}-\d{2}$")

        # 导出：入库期间列也是 YYYY-MM 文本（不是空、也不是完整日期）
        resp2 = self.client.get(
            reverse("export_listing", args=[self.july.id, "purchase_receipt"])
        )
        wb = load_workbook(BytesIO(resp2.content), read_only=True)
        try:
            ws = wb["采购入库"] if "采购入库" in wb.sheetnames else wb[wb.sheetnames[0]]
            values = [r[0] for r in ws.iter_rows(min_row=2, values_only=True)]
        finally:
            wb.close()
        self.assertTrue(values)
        for value in values:
            self.assertRegex(str(value), r"^\d{4}-\d{2}$")

    def test_nav_has_the_entry_and_export_sheet(self):
        html = self.client.get(reverse("period_detail", args=[self.july.id])).content.decode("utf-8")
        nav = html.split('data-nav="import"', 1)[1].split("</div>", 2)[0]
        self.assertIn("采购入库", nav)
        self.assertIn(reverse("purchase_receipts", args=[self.july.id]), nav)
        # 导出《采购入库》
        resp = self.client.get(
            reverse("export_listing", args=[self.july.id, "purchase_receipt"])
        )
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(BytesIO(resp.content), read_only=True)
        try:
            ws = wb["采购入库"] if "采购入库" in wb.sheetnames else wb[wb.sheetnames[0]]
            head = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
            rows = list(ws.iter_rows(min_row=2, values_only=True))
        finally:
            wb.close()
        self.assertEqual(
            head[:5], ["入库期间", "入库日期", "入库单号", "业务类型", "供应商"]
        )
        self.assertIn("累计到票数量", head)
        # 默认（全部期间）导出 2 行；第 3 列＝入库单号（前面是「入库期间」「入库日期」）
        self.assertEqual([r[2] for r in rows], ["GR-RCPT-06", "GR-RCPT-07"])


class OrderCompletionReportPageTests(SettlementFixtureMixin, TestCase):
    """「工单报工」页（用户 2026-09-21 两轮要求）。

    第一轮：①页头六个数字（报工数量 / 完工成本 / 在产品成本 / 总成本 / 报工入库数量 /
      报工入库金额）从卡片改成表格，加「期间」「车间」两列（每个车间一行 + 合计行）；
      ②明细表上方那栏筛选（工单编号 / 车间·部门 / 产品编码）整栏删除；
      ③明细表＝上传的「报工记录」原样 + 6 个计算列，与「导出Excel」逐列一致。
    第二轮：④汇总表把**浆粕车间**也放进去（按浆粕车间自己的流水）；⑤汇总表是**累计**的
      （以后的月份也会加进来），但**默认只显示当月**；⑥汇总表加「期间」「车间」筛选；
      ⑦明细表标题去掉括注，加一列「期间」（＝报工时间的年月）与「期间 / 工单部门 / 工单编号」筛选。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="oc-report-user", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.wo_batch = self.create_batch(
            self.period, ImportBatch.Kind.WORK_ORDER, "oc-wo"
        )
        self.wr_batch = self.create_batch(
            self.period, ImportBatch.Kind.WORK_REPORT, "oc-wr"
        )
        self.wo = WorkOrder.objects.create(
            period=self.period, batch=self.wo_batch, order_no="MO-OC-1",
            production_dept_code="tefang", plan_quantity=Decimal("100"),
        )
        # 另一张工单：既没有报工也没有成本/入库 → 汇总里不该出「染色车间」这一行
        WorkOrder.objects.create(
            period=self.period, batch=self.wo_batch, order_no="MO-OC-2",
            production_dept_code="RanZheng", plan_quantity=Decimal("50"),
        )
        WorkReport.objects.create(
            period=self.period, batch=self.wr_batch, work_order_no="MO-OC-1",
            report_time=timezone.make_aware(datetime(2026, 7, 10, 8, 30)),
            task_no="T-1", operation_no="OP10", operation_name="报工",
            material_code="FG-OC", material_name="示例产品", operator="张三",
            quality_status="合格", quantity=Decimal("10"), uom="kg",
            work_order_department="特纺",
        )
        # 第二条报工：同一个期间、报工时间是 6 月（跨月报工，用来试「期间」筛选）
        WorkReport.objects.create(
            period=self.period, batch=self.wr_batch, work_order_no="MO-OC-1",
            report_time=timezone.make_aware(datetime(2026, 6, 5, 9, 0)),
            task_no="T-2", operation_no="OP20", operation_name="报工",
            material_code="FG-OC", material_name="示例产品", operator="李四",
            quality_status="合格", quantity=Decimal("5"), uom="kg",
            work_order_department="染整（纱线）",
        )
        WorkOrderCostResult.objects.create(
            period=self.period, work_order=self.wo, product_code="FG-OC",
            product_name="示例产品", plan_quantity=Decimal("100"),
            finished_quantity=Decimal("60"), wip_quantity=Decimal("40"),
            finished_unit_cost=Decimal("10"), finished_cost=Decimal("600"),
            wip_cost=Decimal("400"), total_cost=Decimal("1000"),
        )
        product = Material.objects.create(code="FG-OC", name="示例产品", uom="kg")
        txn_batch = self.create_batch(
            self.period, ImportBatch.Kind.INVENTORY, "oc-inv"
        )
        txn = self.create_txn(
            self.period, txn_batch, txn_date=date(2026, 7, 10),
            document_no="MO-OC-1", business_type="报工",
            direction=InventoryTransaction.Direction.IN, material=product,
            quantity=Decimal("12"), order_no="MO-OC-1",
        )
        txn.amount = Decimal("120")
        txn.save(update_fields=["amount"])
        self.url = reverse("order_completion_report", args=[self.period.id])
        self.export_url = reverse(
            "export_listing", args=[self.period.id, "order_completion"]
        )

    # —— 小工具 ——
    @staticmethod
    def _heads(html, css_class):
        block = html.split('class="%s"' % css_class, 1)[1].split("</thead>", 1)[0]
        return [
            re.sub(r"<[^>]+>", "", head).strip()
            for head in re.findall(r"<th[^>]*>(.*?)</th>", block, re.S)
        ]

    @staticmethod
    def _export_head(resp):
        wb = load_workbook(BytesIO(resp.content), read_only=True)
        try:
            ws = wb[wb.sheetnames[0]]
            return [
                "" if cell.value is None else str(cell.value)
                for cell in next(ws.iter_rows(min_row=1, max_row=1))
            ]
        finally:
            wb.close()

    def _summary(self, **params):
        resp = self.client.get(self.url, params)
        return resp, resp.context["summary_rows"], resp.context["totals"]

    def _make_pulp_data(self, *, quantity="10", amount="500"):
        """造一份浆粕车间数据：成品「生产入库」流水 + 原料领料流水 + 车间费用。"""
        self.pulp_batch = self.create_batch(
            self.period, ImportBatch.Kind.PULP_FLOW, "oc-pulp"
        )
        product = Material.objects.create(
            code="21.03", name="芳纶浆粕DFP-103", uom="kg"
        )
        txn = self.create_txn(
            self.period, self.pulp_batch, txn_date=date(2026, 7, 20),
            document_no="PULP-OUT-1", business_type="生产入库",
            direction=InventoryTransaction.Direction.IN, material=product,
            quantity=Decimal(quantity), order_no="PULP202607",
        )
        txn.amount = Decimal(amount)
        txn.save(update_fields=["amount"])
        self.create_txn(
            self.period, self.pulp_batch, txn_date=date(2026, 7, 5),
            document_no="PULP-IN-1", business_type="领料出库",
            direction=InventoryTransaction.Direction.OUT,
            material=Material.objects.create(
                code="R52.101", name="对位芳纶短纤维", uom="kg"
            ),
            quantity=Decimal("8"),
        )
        # 浆粕原料流水（用户上传）：领料出库 + 金额；备注留空 → 按物料名称（芳纶）归到芳纶类成品
        PulpMaterial.objects.create(
            period=self.period, batch=self.pulp_batch, txn_date=date(2026, 7, 5),
            document_no="PULP-IN-1", material_code="R52.101",
            material_name="对位芳纶短纤维", quantity=Decimal("8"),
            unit_price=Decimal("37.5"), amount=Decimal("300"),
            business_type="领料出库", uom="kg",
        )
        WorkshopCostPool.objects.create(
            period=self.period, batch=self.pulp_batch, workshop="浆粕车间",
            water_power=Decimal("100"), wages=Decimal("100"),
        )

    def _make_other_period(self):
        """再建一个期间（2026-08），用来验证汇总表是累计的。"""
        aug = AccountingPeriod.objects.create(month=date(2026, 8, 1))
        wo_batch = self.create_batch(aug, ImportBatch.Kind.WORK_ORDER, "oc-wo-aug")
        wr_batch = self.create_batch(aug, ImportBatch.Kind.WORK_REPORT, "oc-wr-aug")
        order = WorkOrder.objects.create(
            period=aug, batch=wo_batch, order_no="MO-OC-AUG",
            production_dept_code="JiZhi", plan_quantity=Decimal("30"),
        )
        WorkReport.objects.create(
            period=aug, batch=wr_batch, work_order_no="MO-OC-AUG",
            report_time=timezone.make_aware(datetime(2026, 8, 3, 8, 0)),
            task_no="T-8", quantity=Decimal("30"), uom="kg",
            work_order_department="机织",
        )
        WorkOrderCostResult.objects.create(
            period=aug, work_order=order, product_code="FG-AUG",
            product_name="坯布", plan_quantity=Decimal("30"),
            finished_quantity=Decimal("30"), wip_quantity=Decimal("0"),
            finished_cost=Decimal("800"), wip_cost=Decimal("0"),
            total_cost=Decimal("800"),
        )
        return aug

    # —— 汇总表 ——
    def test_summary_is_period_workshop_table(self):
        """六个数字＝带「期间」「车间」的表格（一个车间一行 + 合计行），不再是一排方块。"""
        resp, rows, totals = self._summary()
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        self.assertNotIn('class="summary-strip"', html)
        self.assertEqual(
            self._heads(html, "oc-summary-table"),
            ["期间", "车间", "报工数量", "完工成本", "在产品成本", "总成本",
             "报工入库数量", "报工入库金额"],
        )
        # 全 0 的车间（染色车间）不出行，只剩有数据的纺纱车间
        self.assertEqual([row["workshop_label"] for row in rows], ["纺纱车间"])
        row = rows[0]
        self.assertEqual(row["month_label"], "2026-7")   # 期间格式＝2026-7
        self.assertEqual(row["report_quantity"], Decimal("15"))   # 10 + 5（跨月那条也算本期）
        self.assertEqual(row["finished_cost"], Decimal("600"))
        self.assertEqual(row["wip_cost"], Decimal("400"))
        self.assertEqual(row["total_cost"], Decimal("1000"))
        self.assertEqual(row["receipt_quantity"], Decimal("12"))
        self.assertEqual(row["receipt_amount"], Decimal("120"))
        # 合计＝原来页头六个数字的口径（按本期全部工单统计，不受明细表分页/筛选影响）
        self.assertEqual(totals["report_quantity"], Decimal("15"))
        self.assertEqual(totals["finished_cost"], Decimal("600"))
        self.assertEqual(totals["wip_cost"], Decimal("400"))
        self.assertEqual(totals["total_cost"], Decimal("1000"))
        self.assertEqual(totals["receipt_quantity"], Decimal("12"))
        self.assertEqual(totals["receipt_amount"], Decimal("120"))
        self.assertEqual(resp.context["total_label"], "合计（2026-7 · 1 张工单）")
        footer = html.split("<tfoot>", 1)[1].split("</tfoot>", 1)[0]
        self.assertIn("1,000.00", footer)

    def test_summary_includes_pulp_workshop(self):
        """浆粕车间也要有（按浆粕自己的流水）：数量＝成品入库数量，成本＝浆粕成本归集。"""
        self._make_pulp_data(quantity="10", amount="0")
        resp, rows, totals = self._summary()
        labels = [row["workshop_label"] for row in rows]
        self.assertIn("浆粕车间", labels)
        pulp = [row for row in rows if row["workshop_label"] == "浆粕车间"][0]
        # 报工数量 / 报工入库数量＝浆粕成品「生产入库」流水数量（＝产量）
        self.assertEqual(pulp["report_quantity"], Decimal("10"))
        self.assertEqual(pulp["receipt_quantity"], Decimal("10"))
        # 完工成本＝原料 300 + 水电 100 + 工资 100（浆粕成本归集），在产品为 0
        self.assertEqual(pulp["finished_cost"], Decimal("500"))
        self.assertEqual(pulp["wip_cost"], Decimal("0"))
        self.assertEqual(pulp["total_cost"], Decimal("500"))
        # 报工入库金额取流水上的金额（这里没回填过 → 0）
        self.assertEqual(pulp["receipt_amount"], Decimal("0"))
        # 合计把这些都加上
        self.assertEqual(totals["total_cost"], Decimal("1500"))
        self.assertEqual(totals["finished_cost"], Decimal("1100"))
        self.assertEqual(totals["receipt_quantity"], Decimal("22"))
        # 「车间」筛选里能选到浆粕车间
        self.assertIn(
            "浆粕车间",
            [item["label"] for item in resp.context["workshop_options"]],
        )

    def test_summary_pulp_falls_back_to_flow_amounts(self):
        """浆粕原料流水缺失时，完工成本退回用流水金额（不让汇总页挂掉）。"""
        self._make_pulp_data(quantity="10", amount="500")
        # 删掉原料流水 → collect_pulp_costs 会报错
        PulpMaterial.objects.all().delete()
        _, rows, _ = self._summary()
        pulp = [row for row in rows if row["workshop_label"] == "浆粕车间"][0]
        self.assertEqual(pulp["finished_cost"], Decimal("500"))
        self.assertEqual(pulp["receipt_amount"], Decimal("500"))

    def test_summary_is_cumulative_but_defaults_to_current_month(self):
        """汇总表是累计的（以后的月份会加进来），但默认只显示当月。"""
        self._make_other_period()
        # 默认：只有本期（2026-7）
        resp, rows, _ = self._summary()
        self.assertEqual({row["month_label"] for row in rows}, {"2026-7"})
        self.assertEqual(resp.context["current_month"], "2026-7")
        self.assertIn("2026-8", resp.context["period_options"])
        # 全部期间：两个月都在
        _, all_rows, all_totals = self._summary(s_period="all")
        self.assertEqual(
            {row["month_label"] for row in all_rows}, {"2026-7", "2026-8"}
        )
        self.assertEqual(
            [row["workshop_label"] for row in all_rows if row["month_label"] == "2026-8"],
            ["机织车间"],
        )
        aug = [row for row in all_rows if row["month_label"] == "2026-8"][0]
        self.assertEqual(aug["report_quantity"], Decimal("30"))
        self.assertEqual(aug["total_cost"], Decimal("800"))
        # 累计合计＝两个月相加
        self.assertEqual(all_totals["total_cost"], Decimal("1800"))
        self.assertIn("2 个期间", self._summary(s_period="all")[0].context["total_label"])
        # 指定某个月（新格式 2026-8 与老链接的 2026-08 都要认）
        for value in ("2026-8", "2026-08"):
            _, aug_rows, aug_totals = self._summary(s_period=value)
            self.assertEqual({row["month_label"] for row in aug_rows}, {"2026-8"}, value)
            self.assertEqual(aug_totals["total_cost"], Decimal("800"), value)

    def test_summary_workshop_filter(self):
        """汇总表的「车间」筛选只作用于本表。"""
        _, rows, totals = self._summary(s_workshop="tefang")
        self.assertEqual([row["workshop_label"] for row in rows], ["纺纱车间"])
        self.assertEqual(totals["total_cost"], Decimal("1000"))
        _, none_rows, none_totals = self._summary(s_workshop="JiZhi")
        self.assertEqual(none_rows, [])
        self.assertEqual(none_totals["total_cost"], Decimal("0"))

    # —— 明细表 ——
    def test_detail_table_has_period_column_first(self):
        """明细表第一列＝期间（报工时间的年月），与导出逐列一致。"""
        resp = self.client.get(self.url)
        html = resp.content.decode("utf-8")
        heads = self._heads(html, "oc-detail-table")
        self.assertEqual(heads[0], "期间")
        self.assertEqual(heads[1], "报工时间")
        # 标题里不再有括注，页头那句长说明也删了
        self.assertNotIn("（＝上传的报工记录原样", html)
        self.assertNotIn("明细表＝期间 + 你上传的", html)
        self.assertNotIn("汇总按期间 × 车间累计", html)
        # 每行第一格＝该行报工时间的年月（2026-7 这种格式）
        row = resp.context["page"].object_list[0]
        self.assertEqual(row[0], "2026-6")               # 6 月那条排在前（按 id）
        self.assertEqual(row[1].date(), date(2026, 6, 5))  # 报工时间（本地时间带秒）
        # 与导出完全一致（含列名与列顺序）
        self.assertEqual(
            heads, self._export_head(self.client.get(self.export_url))
        )

    def test_detail_period_filter(self):
        """明细表的「期间」＝报工时间的年月。"""
        both = self.client.get(self.url)
        self.assertEqual(both.context["page"].paginator.count, 2)
        self.assertEqual(both.context["detail_months"], ["2026-6", "2026-7"])
        june = self.client.get(self.url, {"d_period": "2026-6"})
        self.assertEqual(june.context["page"].paginator.count, 1)
        self.assertEqual(june.context["page"].object_list[0][0], "2026-6")
        # 老链接里补零的写法（2026-06 / 2026-07）也照旧认
        self.assertEqual(
            self.client.get(self.url, {"d_period": "2026-06"}).context[
                "page"
            ].paginator.count,
            1,
        )
        july = self.client.get(self.url, {"d_period": "2026-7"})
        self.assertEqual(july.context["page"].paginator.count, 1)
        self.assertEqual(july.context["page"].object_list[0][0], "2026-7")
        # 明细表里的「期间」就是 2026-7 这种格式
        self.assertEqual(
            self._heads(
                self.client.get(self.url).content.decode("utf-8"),
                "oc-detail-table",
            )[0],
            "期间",
        )
        # 导出跟着一起筛（页面与导出一致）
        resp = self.client.get(self.export_url, {"d_period": "2026-7"})
        wb = load_workbook(BytesIO(resp.content), read_only=True)
        try:
            ws = wb[wb.sheetnames[0]]
            rows = list(ws.iter_rows(min_row=2, values_only=True))
        finally:
            wb.close()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][0], "2026-7")

    def test_detail_department_and_order_filters(self):
        """明细表新增的「工单部门」「工单编号」筛选。"""
        resp = self.client.get(self.url, {"d_department": "特纺"})
        self.assertEqual(resp.context["page"].paginator.count, 1)
        self.assertEqual(
            resp.context["page"].object_list[0][1].date(), date(2026, 7, 10)
        )
        self.assertEqual(
            self.client.get(self.url, {"d_department": "染整"}).context[
                "page"
            ].paginator.count,
            1,
        )
        self.assertEqual(
            self.client.get(self.url, {"d_order_no": "MO-OC-1"}).context[
                "page"
            ].paginator.count,
            2,
        )
        self.assertEqual(
            self.client.get(self.url, {"d_order_no": "MO-OC-9"}).context[
                "page"
            ].paginator.count,
            0,
        )
        self.assertEqual(
            resp.context["detail_departments"], ["染整（纱线）", "特纺"]
        )

    def test_filter_bars(self):
        """两个模块各自的筛选栏：明细＝期间/工单部门/工单编号，汇总＝期间/车间。"""
        html = self.client.get(self.url).content.decode("utf-8")
        for name in ("d_period", "d_department", "d_order_no"):
            self.assertIn('name="%s"' % name, html, name)
        for name in ("s_period", "s_workshop"):
            self.assertIn('name="%s"' % name, html, name)
        # 已删除的老筛选栏字段不再出现
        for name in ("order_no", "workshop", "product"):
            self.assertNotIn('name="%s"' % name, html, name)
        self.assertIn(">筛选<", html)
        self.assertIn(">清空<", html)
        self.assertIn("全部期间（累计）", html)
        self.assertIn("全部车间", html)
        # 下拉一改就即时生效（用户 2026-09-21：选了没变化）——4 个下拉都要带 onchange 提交
        for name in ("s_period", "s_workshop", "d_period", "d_department"):
            block = html.split('name="%s"' % name, 1)[1].split("</select>", 1)[0]
            self.assertIn('onchange="this.form.submit()"', block, name)
        # 生效的筛选条件要写在页面上（让人看得见「筛了什么」）
        self.assertIn(
            "本表已按上方条件过滤：车间＝纺纱车间",
            self._summary(s_workshop="tefang")[0].content.decode("utf-8"),
        )
        self.assertIn(
            "本表已按上方条件过滤：期间＝全部期间（累计）",
            self._summary(s_period="all")[0].content.decode("utf-8"),
        )
        detail_html = self.client.get(self.url, {"d_department": "特纺"}).content.decode("utf-8")
        self.assertIn("明细表已按上方条件过滤：工单部门＝特纺", detail_html)
        self.assertIn("共 1 行 / 1 张工单", detail_html)
        # 没筛选时不显示提示；跨模块的隐藏条件只带非空的（URL 干净）
        self.assertNotIn("本表已按上方条件过滤", html)
        self.assertNotIn("明细表已按上方条件过滤", html)
        for name in ("d_period", "d_department", "d_order_no",
                     "s_period", "s_workshop"):
            self.assertNotIn('type="hidden" name="%s"' % name, html, name)
        # 交叉带条件时：明细行里要带上汇总的 s_*，反之亦然（两个模块互不影响）
        with_summary = self.client.get(
            self.url, {"s_workshop": "tefang"}
        ).content.decode("utf-8")
        self.assertIn('type="hidden" name="s_workshop" value="tefang"', with_summary)
        with_detail = self.client.get(
            self.url, {"d_department": "特纺"}
        ).content.decode("utf-8")
        self.assertIn(
            'type="hidden" name="d_department" value="特纺"', with_detail
        )
        # 期间格式＝2026-7（不补零）；老链接补零写法仍能匹配，见 test_detail_period_filter
        self.assertIn('<option value="2026-7"', html)
        self.assertNotIn('<option value="2026-07"', html)
        # 导出Excel / 打印本表已经挪到明细筛选行里，与「筛选」「清空」同一排、底对齐
        form = html.split('<form class="oc-filter-row"', 2)[2].split("</form>", 1)[0]
        for token in ("筛选", "清空", "导出Excel", 'class="print-here"'):
            self.assertIn(token, form, token)
        self.assertLess(form.index("清空"), form.index("导出Excel"))
        self.assertLess(form.index("导出Excel"), form.index('class="print-here"'))
        # 标题行里不再放按钮（导出/打印都在筛选行里）
        heading = html.split("<h2>工单报工明细表</h2>", 1)[1].split("</div>", 2)[0]
        self.assertNotIn("导出Excel", heading)
        # 汇总表的「打印本表」也在它自己的筛选行里（与筛选/清空同一排）；两个标题行都没有按钮
        summary_form = html.split('<form class="oc-filter-row"', 1)[1].split("</form>", 1)[0]
        self.assertIn('class="print-here"', summary_form)
        for block in html.split("<h2>")[1:]:
            self.assertNotIn("heading-actions", block.split("</div>", 2)[0])
        # 口径说明保持简短（长段落已删）
        self.assertNotIn("口径：报工数量＝报工记录的数量合计", html)
        self.assertIn("合计＝当前筛选下的加总", html)
        # 明细筛选只作用于明细：汇总合计不受其影响
        _, rows, totals = self._summary(d_order_no="MO-OC-9")
        self.assertEqual(totals["total_cost"], Decimal("1000"))
        self.assertEqual(len(rows), 1)

    def test_detail_columns_follow_uploaded_file(self):
        """有上传文件时，列名/顺序＝上传的「报工记录」（和「生产工单」「工单投料」一个做法）。"""
        template = (
            Path(settings.BASE_DIR) / "sample_templates" / "报工记录.csv"
        ).read_bytes()
        upload_head = template.decode("utf-8-sig").splitlines()[0].split(",")
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                create_import_batch(
                    period=self.period,
                    kind=ImportBatch.Kind.WORK_REPORT,
                    uploaded_file=SimpleUploadedFile(
                        "报工记录.csv", template, "text/csv"
                    ),
                    user=self.user,
                )
                resp = self.client.get(self.url)
        html = resp.content.decode("utf-8")
        heads = self._heads(html, "oc-detail-table")
        self.assertEqual(len(heads), len(upload_head) + 1 + 6)
        self.assertEqual(heads[0], "期间")
        self.assertEqual(heads[1: 1 + len(upload_head)], upload_head)
        # 列值按上传文件的列顺序排：工单编号 / 报工数量 落在对的那一列
        row = resp.context["page"].object_list[0]
        self.assertEqual(row[heads.index("工单编号")], "MO-OC-1")
        self.assertEqual(row[heads.index("报工数量")], Decimal("5"))
        self.assertEqual(row[heads.index("报工物料编号")], "FG-OC")
        self.assertEqual(row[0], "2026-6")             # 期间＝报工时间的年月
        self.assertEqual(row[-2], Decimal("12"))        # 报工入库数量（按工单汇总）
        self.assertEqual(row[-1], Decimal("120"))       # 报工入库金额

    def test_legacy_query_params_still_work(self):
        """老链接/书签里不带前缀的参数仍然生效（明细表），导出也跟着走。"""
        by_no = self.client.get(self.url, {"order_no": "MO-OC-1"})
        self.assertEqual(by_no.context["page"].paginator.count, 2)
        by_ws = self.client.get(self.url, {"workshop": "tefang"})
        self.assertEqual(by_ws.context["page"].paginator.count, 2)
        by_product = self.client.get(self.url, {"product": "FG-OC"})
        self.assertEqual(by_product.context["page"].paginator.count, 2)
        empty = self.client.get(self.url, {"order_no": "MO-OC-9"})
        self.assertEqual(empty.context["page"].paginator.count, 0)
        # 老参数只筛明细表；汇总表按自己的口径出（本期全部工单）
        self.assertEqual(empty.context["totals"]["total_cost"], Decimal("1000"))
        resp = self.client.get(self.export_url, {"order_no": "MO-OC-9"})
        wb = load_workbook(BytesIO(resp.content), read_only=True)
        try:
            ws = wb[wb.sheetnames[0]]
            self.assertEqual(
                sum(1 for _ in ws.iter_rows(min_row=2, values_only=True)), 0
            )
        finally:
            wb.close()

class InventoryMatchesPageTests(SettlementFixtureMixin, TestCase):
    """「库存流水匹配」页（用户 2026-09-21）：

    ①顶部整块筛选栏（业务日期…匹配状态 + 筛选 + 清空）删除——老链接里带的参数照旧生效；
    ②加一列「期间」＝业务日期的年月（2026-7 这种格式）；
    ③「库存单据」列里的「流水ID xxx」小字删除；
    ④加「导出Excel」，列与页面一致。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="im-page-user", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.batch = self.create_batch(
            self.period, ImportBatch.Kind.INVENTORY, "im-page"
        )
        self.product = Material.objects.create(
            code="RM-IM", name="匹配测试原料", uom="kg"
        )
        # 一条采购入库（会被合同匹配）＋ 一条销售出库（无需匹配）
        self.purchase = self.create_txn(
            self.period, self.batch, txn_date=date(2026, 7, 12),
            document_no="GR-IM-1", business_type="采购入库",
            direction=InventoryTransaction.Direction.IN, material=self.product,
            quantity=Decimal("10"), supplier_code="SUP-IM",
            supplier_name="匹配供应商", order_no="PO-IM",
        )
        self.sales = self.create_txn(
            self.period, self.batch, txn_date=date(2026, 7, 20),
            document_no="SO-IM-1", business_type="销售出库",
            direction=InventoryTransaction.Direction.OUT, material=self.product,
            quantity=Decimal("4"), customer_name="匹配客户",
        )
        contract_batch = self.create_batch(
            self.period, ImportBatch.Kind.PURCHASE_CONTRACT, "im-contract"
        )
        PurchaseContractLine.objects.create(
            period=self.period, batch=contract_batch, contract_no="C-IM-1",
            contract_date=date(2026, 7, 1), supplier_code="SUP-IM",
            supplier_name="匹配供应商", material_code=self.product.code,
            material_name=self.product.name, quantity=Decimal("10"),
            remaining_quantity=Decimal("10"),
            unit_price_excl_tax=Decimal("5"), amount_excl_tax=Decimal("50"),
        )
        auto_match_all(self.period)
        self.url = reverse("inventory_matches", args=[self.period.id])
        self.export_url = reverse(
            "export_listing", args=[self.period.id, "inventory_matches"]
        )

    def test_filter_bar_is_removed(self):
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertIn("<h1>库存流水匹配</h1>", html)
        # 整块筛选栏没了：10 个字段 + 筛选/清空
        for name in ("txn_date", "document_no", "business_type", "direction",
                     "warehouse", "material", "party", "order_no",
                     "match_mode", "match_status"):
            self.assertNotIn('name="%s"' % name, html, name)
        self.assertNotIn('class="filter-card"', html)
        self.assertNotIn(">筛选<", html)
        self.assertNotIn(">清空<", html)
        # 重新自动匹配按钮还在（它不在筛选栏里）
        self.assertIn("重新自动匹配", html)

    def test_legacy_query_params_still_work(self):
        """老链接/书签里带的参数照旧生效（只是页面上看不到筛选栏了），导出也跟着走。"""
        by_type = self.client.get(self.url, {"business_type": "销售出库"})
        self.assertEqual(by_type.context["page"].paginator.count, 1)
        self.assertEqual(
            by_type.context["page"].object_list[0].document_no, "SO-IM-1"
        )
        by_doc = self.client.get(self.url, {"document_no": "GR-IM-1"})
        self.assertEqual(by_doc.context["page"].paginator.count, 1)
        by_party = self.client.get(self.url, {"party": "匹配供应商"})
        self.assertEqual(by_party.context["page"].paginator.count, 1)
        by_date = self.client.get(self.url, {"txn_date": "2026-07-20"})
        self.assertEqual(by_date.context["page"].paginator.count, 1)
        by_mode = self.client.get(self.url, {"match_status": "unmatched"})
        self.assertEqual(by_mode.context["page"].paginator.count, 1)
        # 导出带同样的参数 → 行数与页面一致
        resp = self.client.get(self.export_url, {"business_type": "销售出库"})
        wb = load_workbook(BytesIO(resp.content), read_only=True)
        try:
            ws = wb[wb.sheetnames[0]]
            rows = list(ws.iter_rows(min_row=2, values_only=True))
        finally:
            wb.close()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][1], "SO-IM-1")

    def test_period_column_and_no_flow_id(self):
        """加了「期间」列（业务日期的年月）；库存单据列里的「流水ID」小字删了。"""
        resp = self.client.get(self.url)
        html = resp.content.decode("utf-8")
        heads = [re.sub(r"<[^>]+>", "", x).strip()
                 for x in re.findall(r"<th[^>]*>(.*?)</th>", html, re.S)]
        self.assertEqual(heads[0], "期间")
        self.assertEqual(heads[1], "库存单据")
        self.assertNotIn("流水ID", html)
        rows = {txn.document_no: txn for txn in resp.context["page"].object_list}
        self.assertEqual(rows["GR-IM-1"].period_label, "2026-7")
        self.assertEqual(rows["SO-IM-1"].period_label, "2026-7")
        # 页面上确实按 2026-7 显示（不是 2026-07）
        first_row = html.split("<tbody>", 1)[1].split("</tr>", 1)[0]
        self.assertIn(">2026-7<", first_row)
        self.assertNotIn(">2026-07<", first_row)
        # 每行 16 格（加期间列后）
        body = html.split("<tbody>", 1)[1].split("</tbody>", 1)[0]
        for tr in re.findall(r"<tr\b[^>]*>(.*?)</tr>", body, re.S):
            self.assertEqual(len(re.findall(r"<td[^>]*>", tr)), 16, tr[:60])

    def test_export_columns_match_page(self):
        """导出Excel：列与页面一致（期间 / 匹配到的单据 / 处理方式 / 匹配状态）。"""
        resp = self.client.get(self.export_url)
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(BytesIO(resp.content), read_only=True)
        try:
            ws = wb[wb.sheetnames[0]]
            heads = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
            rows = list(ws.iter_rows(min_row=2, values_only=True))
        finally:
            wb.close()
        self.assertEqual(
            heads,
            ["期间", "库存单据", "业务日期", "匹配类别", "业务类型", "方向", "仓库",
             "物料编码", "物料名称", "数量", "供应商/客户", "订单号",
             "匹配到的单据", "处理方式", "匹配状态"],
        )
        self.assertEqual(len(rows), 2)
        by_doc = {row[1]: row for row in rows}
        purchase = by_doc["GR-IM-1"]
        self.assertEqual(purchase[0], "2026-7")          # 期间＝业务日期的年月
        self.assertEqual(purchase[3], "采购入库")         # 匹配类别
        self.assertIn("合同 C-IM-1", purchase[12])        # 匹配到的单据
        self.assertEqual(purchase[13], "自动匹配")        # 处理方式
        self.assertEqual(purchase[14], "已匹配")          # 匹配状态
        self.assertEqual(by_doc["SO-IM-1"][3], "无需匹配")  # 销售出库不参与匹配

    def test_export_center_has_entry(self):
        resp = self.client.get(reverse("export_center", args=[self.period.id]))
        self.assertContains(resp, "库存流水匹配")
        self.assertContains(resp, 'value="inventory_matches"')

class WorkshopCostPageLayoutTests(SettlementFixtureMixin, TestCase):
    """车间费用页（用户 2026-09-21）：

    ①录入表数据带分节号；②「导出Excel」放在「打印本表」旁边（录入表 + 汇总表各一处，页头不再有）；
    ③车间字段的「原上传」小字删除；④「说明」精简；⑤汇总表新增专用导出（含总合计，且总合计不为 0）。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="wc-layout", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        batch = self.create_batch(self.period, ImportBatch.Kind.WORKSHOP_COST, "wc-layout")
        # 两个车间，数值够大以便检查千分位
        WorkshopCostPool.objects.create(
            period=self.period, batch=batch, workshop="RanZheng",
            wages=Decimal("100723.08"), depreciation=Decimal("91060.70"),
            water_power=Decimal("41553.11"),
        )
        WorkshopCostPool.objects.create(
            period=self.period, batch=batch, workshop="tefang",
            wages=Decimal("831965.69"), depreciation=Decimal("703137.12"),
            consumables=Decimal("198446.17"),
        )
        self.url = reverse("workshop_costs", args=[self.period.id])

    def test_entry_values_have_thousand_separators(self):
        """录入表数据带分节号（1,234.56）；保存时前端/后端都会去掉逗号。"""
        resp = self.client.get(self.url)
        html = resp.content.decode("utf-8")
        rows = {row["workshop"]: row for row in resp.context["entry_group"]["rows"]}
        cells = {c["field"]: c for c in rows["RanZheng"]["cells"]}
        self.assertEqual(cells["wages"]["input_value"], "100,723.08")
        self.assertEqual(cells["depreciation"]["input_value"], "91,060.70")
        self.assertEqual(cells["water_power"]["input_value"], "41,553.11")
        self.assertIn('value="100,723.08"', html)
        # 合计列本来就有分节号
        # 行合计＝100723.08+91060.70+41553.11＝233,336.89（录入表）
        self.assertIn("233,336.89", html)
        # 带分节号提交也能存进去（后端会去掉逗号）
        data = {"csrfmiddlewaretoken": "x"}
        for field, _label in views.WORKSHOP_COST_COLUMNS:
            data["ws_%s_RanZheng_%s" % (self.period.id, field)] = "0"
        data["ws_%s_RanZheng_wages" % self.period.id] = "1,234.56"
        self.client.post(self.url, data, follow=True)
        pool = WorkshopCostPool.objects.get(period=self.period, workshop="RanZheng")
        self.assertEqual(pool.wages, Decimal("1234.56"))

    def test_export_buttons_sit_next_to_print(self):
        """导出Excel 在「打印本表」左边，且录入表/汇总表各一处；页头不再有导出。"""
        html = self.client.get(self.url).content.decode("utf-8")
        page_head = html.split("</section>", 1)[0]
        self.assertNotIn("导出Excel", page_head)
        for anchor in ("各车间费用（", "车间费用汇总表（按车间）"):
            seg = html.split(anchor, 1)[1][:1500]
            self.assertIn("导出Excel", seg, anchor)
            self.assertIn('class="print-here"', seg, anchor)
            self.assertLess(
                seg.index("导出Excel"), seg.index('class="print-here"'), anchor
            )
        # 两个导出各自指向对的 kind
        self.assertIn("export-list/workshop_cost/", html)
        self.assertIn("export-list/workshop_cost_summary/", html)

    def test_summary_table_filter_by_period_and_workshop(self):
        """汇总表自己的筛选（用户 2026-09-27）：期间 + 车间；
        按钮组靠左（align-start）；只影响本表与它的「导出Excel」。"""
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertIn('name="sum_period"', html)
        self.assertIn('name="sum_workshop"', html)
        self.assertIn("filter-actions align-start", html)
        self.assertIn("导出Excel", html)

        # 不筛选：染色 + 纺纱两个车间都在
        resp = self.client.get(self.url)
        labels = [row["label"] for row in resp.context["summary_rows"]]
        self.assertIn("染色车间", labels)
        self.assertIn("纺纱车间", labels)

        # 车间筛选：只剩染色，且合计跟着走
        only_dye = self.client.get(self.url, {"sum_workshop": "RanZheng"})
        dye_rows = only_dye.context["summary_rows"]
        self.assertEqual([r["label"] for r in dye_rows], ["染色车间"])
        self.assertEqual(
            only_dye.context["summary_footer"]["total_text"], dye_rows[0]["total_text"]
        )
        self.assertTrue(only_dye.context["summary_filter"]["active"])

        # 期间筛选：7 月有数据；6 月没有 → 空表
        july = self.client.get(self.url, {"sum_period": "2026-07"})
        self.assertEqual(len(july.context["summary_rows"]), len(resp.context["summary_rows"]))
        june = self.client.get(self.url, {"sum_period": "2026-06"})
        self.assertEqual(june.context["summary_rows"], [])
        self.assertIsNone(june.context["summary_footer"])

        # 导出跟随筛选：只剩染色一行（+ 末行总合计）
        exp = self.client.get(
            reverse("export_listing", args=[self.period.id, "workshop_cost_summary"]),
            {"sum_workshop": "RanZheng"},
        )
        wb = load_workbook(BytesIO(exp.content), read_only=True)
        try:
            exported = list(wb[wb.sheetnames[0]].iter_rows(values_only=True))
        finally:
            wb.close()
        body = [row for row in exported[1:] if row[2] and row[2] != "总合计"]
        self.assertEqual(len(body), 1)
        self.assertEqual(body[0][2], "染色车间")

    def test_uploaded_marker_and_note_are_gone(self):
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertNotIn("原上传", html)
        # 说明精简成 3 条
        note = html.split('class="eyebrow">HOW IT IS USED', 1)[1]
        note = note.split("</ul>", 1)[0]
        self.assertEqual(note.count("<li>"), 3)
        self.assertNotIn("写进对应期间的车间费用池", note)
        self.assertIn("加工费", note)

    def test_summary_export_has_real_total(self):
        """汇总导出：按车间累计 + 末行总合计（总合计曾因千分位被算成 0）。"""
        resp = self.client.get(
            reverse("export_listing", args=[self.period.id, "workshop_cost_summary"])
        )
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(BytesIO(resp.content), read_only=True)
        try:
            ws = wb[wb.sheetnames[0]]
            rows = list(ws.iter_rows(values_only=True))
        finally:
            wb.close()
        heads = list(rows[0])
        self.assertEqual(
            heads,
            ["顺序", "期间", "车间", "人工费", "折旧费", "水电费", "蒸汽费",
             "燃气费", "染料助剂费", "机物料费", "加工费", "合计"],
        )
        body = {row[2]: row for row in rows[1:]}
        self.assertEqual(
            Decimal(str(body["染色车间"][3])), Decimal("100723.08")
        )
        self.assertEqual(
            Decimal(str(body["纺纱车间"][3])), Decimal("831965.69")
        )
        total = body["总合计"]
        self.assertEqual(
            Decimal(str(total[3])), Decimal("100723.08") + Decimal("831965.69")
        )
        self.assertNotEqual(Decimal(total[-1]), Decimal("0"))   # 总合计不再为 0
        # 页面上的总合计行也是真数
        html = self.client.get(self.url).content.decode("utf-8")
        footer = html.split("<tfoot>", 1)[1].split("</tfoot>", 1)[0]
        self.assertIn("932,688.77", footer)      # 100723.08 + 831965.69

class WorkshopSettlementOrderLayoutTests(SettlementFixtureMixin, TestCase):
    """车间独立结转页（用户 2026-09-21）：

    ①序号方框放大（原来被 app.css 的 td input{height:17px} 压到 17px，太小不协调）；
    ②「完整结算按钮的行为…」不再是单独一个小模块，挪到表格左下角（与状态提示同一列）；
    ③「保存结转顺序」放右下角；
    ④状态提示（已有车间完成结转…／已锁账…）紧挨着说明文字。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="ws-order-layout", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.url = reverse("workshop_settlements", args=[self.period.id])

    def _html(self):
        return self.client.get(self.url).content.decode("utf-8")

    def _make_workshops(self, period=None):
        """造两个车间（页面按「生产工单的生产部门编号」识别车间）。"""
        target = period or self.period
        batch = self.create_batch(target, ImportBatch.Kind.WORK_ORDER, "ws-order")
        for code in ("RanZheng", "tefang"):
            WorkOrder.objects.create(
                period=target, batch=batch, order_no="MO-%s" % code,
                production_dept_code=code, plan_quantity=Decimal("10"),
            )

    def test_sequence_box_is_enlarged(self):
        # 需要先有车间行，页面才会渲染出结转顺序表
        self._make_workshops()
        html = self._html()
        # 页面自带样式里显式放大（覆盖 app.css 里 17px 的那条）
        self.assertIn(".sequence-input", html)
        seg = html.split(".sequence-input", 1)[1].split("}", 1)[0]
        self.assertIn("height: 36px", seg)
        self.assertIn("width: 92px", seg)
        self.assertIn("font-size: 16px", seg)
        # 这段页面样式必须排在 app.css 之后，否则覆盖不了
        self.assertGreater(
            html.index(".sequence-input"),
            html.index("settlement/app."),
        )
        # 编号格仍在表格里（结构没变）
        self.assertIn('class="sequence-input"', html)

    def test_full_settlement_note_is_not_a_standalone_block(self):
        """说明不再是单独一个小模块；整页只保留「结转顺序」这一个 section-card。"""
        html = self._html()
        self.assertEqual(html.count('<section class="section-card"'), 1)
        self.assertIn("完整结算按钮的行为", html)
        # 说明落在表尾左侧那块（settle-footer-notes）里
        footer = html.split('class="settle-footer"', 1)[1]
        notes = footer.split('class="settle-footer-notes"', 1)[1].split("</div>", 1)[0]
        self.assertIn("完整结算按钮的行为", notes)
        self.assertIn("自动继续处理所有“待结转”车间", notes)
        self.assertIn("已经独立结转的车间不会重复计算", notes)

    def test_save_button_is_bottom_right(self):
        html = self._html()
        footer = html.split('class="settle-footer"', 1)[1]
        self.assertIn('class="settle-footer-action"', footer)
        action = footer.split('class="settle-footer-action"', 1)[1]
        self.assertIn("保存结转顺序", action)
        # 说明在按钮之前出现（左侧），按钮在右侧那块里
        self.assertLess(
            html.index("完整结算按钮的行为"), html.index("settle-footer-action")
        )

    def test_state_hint_sits_next_to_the_note(self):
        # 未结转：提示在说明同一列；置灰按钮也在表尾右侧
        self._make_workshops()
        html = self._html()
        notes = html.split('class="settle-footer-notes"', 1)[1].split("</div>", 1)[0]
        self.assertIn("在上表填写顺序后点此保存", notes)
        # 把一个车间标成已结转：提示换成「顺序已锁定」且仍在同一列
        settlement = WorkshopSettlement.objects.get(
            period=self.period, workshop="RanZheng"
        )
        settlement.status = WorkshopSettlement.Status.SETTLED
        settlement.save(update_fields=["status"])
        html2 = self._html()
        notes2 = html2.split('class="settle-footer-notes"', 1)[1].split("</div>", 1)[0]
        self.assertIn("已有车间完成结转，顺序已锁定", notes2)
        self.assertIn("请先点右上角「重置车间结转」", notes2)
        # 按钮置灰仍在右下角
        action2 = html2.split('class="settle-footer-action"', 1)[1]
        self.assertIn("disabled", action2)

    def test_no_workshop_branch_keeps_same_layout(self):
        """没有车间时：说明与按钮仍在同一个表尾（左说明、右按钮）。"""
        empty_period = AccountingPeriod.objects.create(month=date(2026, 9, 1))
        # 这个期间没有任何工单/车间费用 → 页面走「尚未识别到车间」分支
        html = self.client.get(
            reverse("workshop_settlements", args=[empty_period.id])
        ).content.decode("utf-8")
        self.assertIn("尚未识别到车间", html)
        self.assertIn('class="settle-footer"', html)
        notes = html.split('class="settle-footer-notes"', 1)[1].split("</div>", 1)[0]
        self.assertIn("完整结算按钮的行为", notes)
        self.assertIn("尚未识别到车间，暂无可保存的顺序", notes)
        action = html.split('class="settle-footer-action"', 1)[1]
        self.assertIn("保存结转顺序", action)

class InventoryResultsPageTests(SettlementFixtureMixin, TestCase):
    """「库存金额结转」页（用户 2026-09-21）：

    ①两个表各自带「导出Excel」（放在表标题旁、打印本表左边），页头不再放导出；
    ②删掉「金额单位：元；期间为本次核算期间。」与「共 N 个物料」；
    ③汇总加筛选「期间 / 物料分类」、明细加筛选「期间 / 物料编码」——默认只显示当月，可选历史/全部；
    ④明细加一列「期间」；
    ⑤新增「库存金额结转汇总（按物料分类）」导出。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="ir-page", password="test-password"
        )
        self.client.force_login(self.user)
        self.july = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.june = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        self.batch = self.create_batch(
            self.july, ImportBatch.Kind.OPENING_INVENTORY, "ir-page"
        )
        # 本期（7 月）两个物料，分属两个分类
        self._valuation(self.july, "11.01", "纱线")
        self._valuation(self.july, "R21.104", "原料", oq="5", oa="50")
        # 历史期间（6 月）一个物料：用来验证「默认当月、可选历史」
        self._valuation(self.june, "9.99", "面料", oq="7", oa="70")
        self.url = reverse("inventory_results", args=[self.july.id])

    def _valuation(self, period, code, category, *, oq="10", oa="100"):
        OpeningInventoryBalance.objects.create(
            period=period,
            batch=self.batch if period.id == self.july.id else self.create_batch(
                period, ImportBatch.Kind.OPENING_INVENTORY, "ir-%s" % code
            ),
            material_code=code, material_name="物料%s" % code,
            category_name=category, quantity=Decimal("1"),
            amount=Decimal("1"), unit_price=Decimal("1"),
        )
        return InventoryValuation.objects.create(
            period=period, material_code=code, material_name="物料%s" % code,
            opening_quantity=Decimal(oq), opening_amount=Decimal(oa),
            receipt_quantity=Decimal("0"), receipt_amount=Decimal("0"),
            issue_quantity=Decimal("0"), issue_amount=Decimal("0"),
            weighted_average_price=Decimal("10"),
            ending_quantity=Decimal(oq), ending_amount=Decimal(oa),
            converged=True,
        )

    def test_export_buttons_are_in_the_filter_row(self):
        """导出Excel 与「打印本表」跟「筛选/清空」同一排（用户 2026-09-21）。"""
        html = self.client.get(self.url).content.decode("utf-8")
        page_head = html.split("</section>", 1)[0]
        self.assertNotIn("导出Excel", page_head)      # 页头不再有导出
        # 两处筛选行：每处都含 筛选/清空/导出Excel/打印本表，且顺序为 筛选→清空→导出→打印
        forms = html.split('<form class="ir-filter-row"')[1:]
        self.assertEqual(len(forms), 2)
        for form in forms:
            body = form.split("</form>", 1)[0]
            for token in (">筛选<", ">清空<", "导出Excel", 'class="print-here"'):
                self.assertIn(token, body, token)
            self.assertLess(body.index(">清空<"), body.index("导出Excel"))
            self.assertLess(body.index("导出Excel"), body.index('class="print-here"'))
        # 标题行里不再挂按钮（只有页面标题那一条 heading-actions 保留）
        for anchor in ("期末库存汇总（按物料分类）", "期末库存（平台计价）"):
            title_row = html.split(anchor, 1)[1].split("</div>", 2)[0]
            self.assertNotIn("导出Excel", title_row, anchor)
        # 页面样式里说明这四个按钮同高、底对齐
        style = html.split("<style>", 1)[1].split("</style>", 1)[0]
        self.assertIn(".ir-filter-row .filter-actions .button", style)
        self.assertIn("min-height:34px", style)
        self.assertIn("align-self:flex-end", style)
        self.assertIn("export-list/inventory_valuation_summary/", html)
        self.assertIn("export-list/inventory_valuation/", html)

    def test_removed_texts_are_gone(self):
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertNotIn("金额单位：元", html)
        self.assertNotIn("期间为本次核算期间", html)
        self.assertIsNone(re.search(r"共\s*[\d,]+\s*个物料", html))

    def test_filters_default_to_current_month(self):
        """默认只显示当月；选「全部期间」能看到历史（6 月）。"""
        html = self.client.get(self.url).content.decode("utf-8")
        for name in ("s_period", "s_category", "d_period", "d_material"):
            self.assertIn('name="%s"' % name, html, name)
        # 默认：只有 7 月的行
        resp = self.client.get(self.url)
        self.assertEqual(
            {row["period_label"] for row in resp.context["category_rows"]}, {"2026-07"}
        )
        self.assertEqual(
            {item.material_code for item in resp.context["page"].object_list},
            {"11.01", "R21.104"},
        )
        # 明细：全部期间 → 6 月的物料也出来
        allresp = self.client.get(self.url, {"d_period": "all"})
        self.assertIn(
            "9.99", {i.material_code for i in allresp.context["page"].object_list}
        )
        # 汇总：全部期间 → 6 月的分类行也出来（汇总用的是 s_period）
        sum_all = self.client.get(self.url, {"s_period": "all"})
        self.assertIn(
            "2026-06",
            {row["period_label"] for row in sum_all.context["category_rows"]},
        )
        self.assertEqual(
            {row["period_label"] for row in self.client.get(self.url).context["category_rows"]},
            {"2026-07"},
        )
        # 指定 6 月（明细 / 汇总各自）
        june = self.client.get(self.url, {"d_period": "2026-06"})
        self.assertEqual(
            {i.material_code for i in june.context["page"].object_list}, {"9.99"}
        )
        june_sum = self.client.get(self.url, {"s_period": "2026-06"})
        self.assertEqual(
            {row["period_label"] for row in june_sum.context["category_rows"]},
            {"2026-06"},
        )

    def test_material_and_category_filters(self):
        by_code = self.client.get(self.url, {"d_material": "11.01"})
        self.assertEqual(
            {i.material_code for i in by_code.context["page"].object_list}, {"11.01"}
        )
        by_cat = self.client.get(self.url, {"s_category": "原料"})
        self.assertEqual(
            [row["category"] for row in by_cat.context["category_rows"]], ["原料"]
        )
        by_cat2 = self.client.get(self.url, {"s_category": "纱线"})
        self.assertEqual(
            [row["category"] for row in by_cat2.context["category_rows"]], ["纱线"]
        )

    def test_detail_table_has_period_column(self):
        resp = self.client.get(self.url)
        html = resp.content.decode("utf-8")
        seg = html.split("期末库存（平台计价）", 1)[1]
        heads = [re.sub(r"<[^>]+>", "", x).strip()
                 for x in re.findall(r"<th[^>]*>(.*?)</th>", seg.split("</thead>", 1)[0], re.S)]
        self.assertEqual(heads[0], "期间")
        self.assertEqual(heads[1], "物料编码")
        body = seg.split("<tbody>", 1)[1].split("</tbody>", 1)[0]
        for tr in re.findall(r"<tr\b[^>]*>(.*?)</tr>", body, re.S):
            self.assertEqual(len(re.findall(r"<td[^>]*>", tr)), len(heads))
        first = re.findall(r"<td[^>]*>(.*?)</td>", body, re.S)[0]
        self.assertEqual(re.sub(r"<[^>]+>", "", first).strip(), "2026-07")
        # 期间下拉里不应重复出现本期
        self.assertEqual(html.count('<option value="2026-07"'), 0)

    def test_summary_export_columns_and_totals(self):
        resp = self.client.get(
            reverse(
                "export_listing",
                args=[self.july.id, "inventory_valuation_summary"],
            )
        )
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(BytesIO(resp.content), read_only=True)
        try:
            ws = wb[wb.sheetnames[0]]
            rows = list(ws.iter_rows(values_only=True))
        finally:
            wb.close()
        self.assertEqual(
            list(rows[0]),
            ["期间", "物料分类", "期初数量", "期初金额", "入库数量", "入库金额",
             "出库数量", "出库金额", "期末数量", "期末金额"],
        )
        body = {row[1]: row for row in rows[1:]}
        self.assertEqual(Decimal(str(body["原料"][3])), Decimal("50"))
        self.assertEqual(Decimal(str(body["纱线"][3])), Decimal("100"))
        self.assertEqual(rows[-1][0], "合计")
        self.assertEqual(Decimal(str(rows[-1][3])), Decimal("150"))
        # 汇总导出也跟筛选走
        resp2 = self.client.get(
            reverse(
                "export_listing",
                args=[self.july.id, "inventory_valuation_summary"],
            ),
            {"s_category": "原料"},
        )
        wb = load_workbook(BytesIO(resp2.content), read_only=True)
        try:
            ws = wb[wb.sheetnames[0]]
            rows2 = list(ws.iter_rows(values_only=True))
        finally:
            wb.close()
        self.assertEqual(len(rows2), 3)          # 表头 + 原料 + 合计
        self.assertEqual(rows2[1][1], "原料")

    def test_detail_export_has_period_column_and_filters(self):
        resp = self.client.get(
            reverse("export_listing", args=[self.july.id, "inventory_valuation"]),
            {"d_material": "11.01"},
        )
        wb = load_workbook(BytesIO(resp.content), read_only=True)
        try:
            ws = wb[wb.sheetnames[0]]
            rows = list(ws.iter_rows(values_only=True))
        finally:
            wb.close()
        self.assertEqual(list(rows[0])[:4], ["期间", "物料编码", "物料名称", "物料分类"])
        self.assertEqual(len(rows), 2)           # 表头 + 1 行
        self.assertEqual(rows[1][0], "2026-07")
        self.assertEqual(rows[1][1], "11.01")

    def test_export_center_has_summary_entry(self):
        resp = self.client.get(reverse("export_center", args=[self.july.id]))
        self.assertContains(resp, "期末库存汇总（按物料分类，平台计价）")
        self.assertContains(resp, 'value="inventory_valuation_summary"')


class UnitCostDiffReasonTests(TestCase):
    """单位成本差异原因（用户 2026-09-30）：|差异率| ≥ 30% 才给原因，且要说真实原因。"""

    @staticmethod
    def _item(**kwargs):
        from types import SimpleNamespace

        base = dict(
            unit_cost_diff_pct=Decimal("-42.09"),
            finished_quantity=Decimal("832.8"),
            plan_quantity=Decimal("800"),
            direct_material=Decimal("17224.59"),
            wip_quantity=Decimal("0"),
            accum_finished=Decimal("832.8"),
            finished_unit_cost=Decimal("25.76"),
            standard_unit_cost=Decimal("44.48"),
        )
        base.update(kwargs)
        return SimpleNamespace(**base)

    def test_below_standard_due_to_missing_material(self):
        """31.11-YT-1456013：BOM 两种料只投了一种 → 原因写清少投了哪种料。"""
        from settlement.services.workorder_costing import unit_cost_diff_reason

        reason = unit_cost_diff_reason(
            self._item(),
            Decimal("34.0"),
            [("11.716", "对位芳纶纱线20S/2")],
        )
        self.assertIn("少投料", reason)
        self.assertIn("11.716", reason)

    def test_no_reason_within_threshold(self):
        """|差异率| < 30% 不给原因（列里显示 —）。"""
        from settlement.services.workorder_costing import unit_cost_diff_reason

        self.assertEqual(
            unit_cost_diff_reason(self._item(unit_cost_diff_pct=Decimal("-12.5")),
                                  Decimal("34.0"), []),
            "",
        )
        self.assertEqual(
            unit_cost_diff_reason(self._item(unit_cost_diff_pct=Decimal("29.9")),
                                  Decimal("34.0"), []),
            "",
        )

    def test_above_standard_reasons(self):
        """高于标准：投料超 BOM / 工资制造费用高于标准。"""
        from settlement.services.workorder_costing import unit_cost_diff_reason

        over_material = unit_cost_diff_reason(
            self._item(unit_cost_diff_pct=Decimal("55"), direct_material=Decimal("50000")),
            Decimal("34.0"), [],
        )
        self.assertIn("投料超 BOM", over_material)
        overhead = unit_cost_diff_reason(
            self._item(unit_cost_diff_pct=Decimal("99.19"),
                       direct_material=Decimal("20.80"),
                       finished_quantity=Decimal("1.6"),
                       plan_quantity=Decimal("1.6")),
            Decimal("12.9991"), [],
        )
        self.assertIn("工资/制造费用", overhead)


class ProductionCostMergeTests(SettlementFixtureMixin, TestCase):
    """「生产成本」与「总成本表」合并后的页面（用户 2026-09-21）。

    界面要求：无「查看范围 / 六个数 / 生产成本闭合 / 本期费用构成」；
    车间生产成本汇总 = 车间/工单数/期初在制/本期材料/本期费用/本期生产投入/完工成本/期末在制/闭合差额
    ＋ 合计行，并带自己的「期间 + 车间」筛选；
    成本项目×四大块 = 原总成本表矩阵，也带自己的「期间 + 车间」筛选，勾稽胶囊在**右下角**；
    两张表各有一段**文字**口径说明（不是四列表）。
    旧地址 /total-cost/ 跳到新页；导出一个按钮导两张表。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="pc-merge", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        batch = self.create_batch(self.period, ImportBatch.Kind.WORK_ORDER, "pc-merge")
        for code in ("RanZheng", "tefang"):
            order = WorkOrder.objects.create(
                period=self.period, batch=batch, order_no="MO-%s" % code,
                production_dept_code=code, plan_quantity=Decimal("100"),
            )
            WorkOrderCostResult.objects.create(
                period=self.period, work_order=order, product_code="P-%s" % code,
                product_name="产品%s" % code, plan_quantity=Decimal("100"),
                finished_quantity=Decimal("80"), wip_quantity=Decimal("20"),
                direct_material=Decimal("1000"), water_power=Decimal("100"),
                wages=Decimal("200"), finished_material=Decimal("800"),
                finished_wages=Decimal("160"), finished_overhead=Decimal("80"),
                wip_material=Decimal("200"), wip_wages=Decimal("40"),
                wip_overhead=Decimal("20"), finished_cost=Decimal("1040"),
                wip_cost=Decimal("260"), total_cost=Decimal("1300"),
            )
        WorkshopSettlement.objects.create(
            period=self.period, workshop="RanZheng", sequence=1,
            status=WorkshopSettlement.Status.SETTLED,
        )
        self.url = reverse("production_cost_summary", args=[self.period.id])

    @staticmethod
    def _heads(html, anchor, occurrence=1):
        # 取第 occurrence 次出现「之后」那一段（split 的 [0] 是锚点之前的）
        seg = html.split(anchor, occurrence)[occurrence]
        return [re.sub(r"<[^>]+>", "", x).strip()
                for x in re.findall(r"<th[^>]*>(.*?)</th>",
                                    seg.split("</thead>", 1)[0], re.S)]

    def test_case_summary_lists_unpriced_bom_orders(self):
        """用户 2026-09-30：⑥「BOM 子项缺单价的工单」是**第 6 种情形**——

        小结里和 ①~⑤ 一样是蓝色可点筛选的一行、计入合计；明细表的情形也能筛到它。
        """
        batch = self.create_batch(self.period, ImportBatch.Kind.WORK_ORDER, "no-price")
        BomLine.objects.create(
            period=self.period, batch=batch, work_order_no="MO-tefang",
            material_code="P-X", sub_item_code="R-NO-PRICE", sub_item_name="没价的料",
            qty_numerator=Decimal("1"), qty_denominator=Decimal("1"),
        )
        resp = self.client.get(self.url)
        summary = resp.context["issue_case_summary"]
        # ⑦ 已完工/基本完工 + ⑥ 缺单价 都是常驻行；⑥ 按 case 定位
        unpriced = next(item for item in summary if item["case"] == CASE_NO_PRICE)
        self.assertEqual(unpriced["count"], 1)                  # 只有 MO-tefang 的 BOM 缺价
        self.assertIn("R-NO-PRICE", unpriced["basis"])          # 缺的物料编码列出来
        self.assertTrue(unpriced["case_link"])                  # 和上面一样可点筛选（蓝色链接）
        # 计入合计：合计＝明细行数（⑥ 已是独立情形，不与 ①~⑤ 重叠）
        self.assertEqual(
            sum(item["count"] for item in summary), len(resp.context["issue_rows"])
        )
        html = resp.content.decode("utf-8")
        self.assertIn("⑥ BOM 子项缺单价的工单", html)
        self.assertNotIn("不计入合计", html)
        # 明细行：情形＝⑥，提示里也带缺价物料
        row = next(r for r in resp.context["issue_rows"] if r["order_no"] == "MO-tefang")
        self.assertEqual(row["case"], CASE_NO_PRICE)
        self.assertIn("R-NO-PRICE", row["flags"])
        # 情形筛选里 6 种都能选到（①②③④ + ⑤ 正常 + ⑥ 缺单价）
        choices = [c["value"] for c in resp.context["issue_case_choices"]]
        self.assertEqual(
            choices, list(CASE_CHOICES) + [CASE_OTHER, CASE_NO_PRICE]
        )
        # 按 ⑥ 筛选：明细只剩那张缺价单
        filtered = self.client.get(self.url, {"i_case": CASE_NO_PRICE})
        self.assertEqual(
            [r["order_no"] for r in filtered.context["issue_rows"]], ["MO-tefang"]
        )
        # 按 ⑤ 筛选也能命中（明细行上是"— 正常"）
        filtered = self.client.get(self.url, {"i_case": CASE_OTHER})
        self.assertEqual(
            [r["order_no"] for r in filtered.context["issue_rows"]], ["MO-RanZheng"]
        )

    def test_removed_modules_are_gone(self):
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertIn("<h1>生产成本</h1>", html)
        # 页头那句「按车间看本期材料投入…」已按用户要求删除
        self.assertNotIn("按科目列出", html)
        self.assertNotIn("按车间看本期材料投入", html)
        # ① 查看范围字段 ② 六个数 ③ 生产成本闭合 ⑦ 本期费用构成 都不在了
        self.assertNotIn("查看范围", html)
        self.assertNotIn("生产成本闭合", html)
        self.assertNotIn("本期费用构成", html)
        self.assertNotIn('class="cost-equation"', html)
        self.assertNotIn('class="summary-strip', html)
        for label in ("期初在制", "本期材料"):
            # 这几个词现在只出现在表格表头里，不再有独立的数字模块
            self.assertIn(label, html, label)

    def test_summary_table_columns_total_row_and_filters(self):
        html = self.client.get(self.url).content.decode("utf-8")
        heads = self._heads(html, "车间生产成本汇总")
        self.assertEqual(
            heads,
            ["车间", "工单数", "期初在制", "本期材料", "本期费用", "本期生产投入",
             "完工成本", "期末在制", "闭合差额"],
        )
        self.assertNotIn("结转状态", html)
        # 车间列下的「完工数量 / 在制数量」小字已删（只看这张表；速查表的建议文案里会出现该词）
        summary_seg = html.split("车间生产成本汇总", 1)[1].split("</table>", 1)[0]
        self.assertNotIn("完工数量", summary_seg)
        # 合计行
        self.assertIn("<tfoot>", html)
        seg = html.split("车间生产成本汇总", 1)[1].split("</table>", 1)[0]
        foot = seg.split("<tfoot>", 1)[1]
        self.assertIn("合计", foot)
        # 自己的筛选：期间 + 车间
        form = html.split("车间生产成本汇总", 1)[1].split("</form>", 1)[0]
        self.assertIn('name="p_period"', form)
        self.assertIn('name="p_workshop"', form)
        self.assertIn(">筛选<", form)
        self.assertIn(">清空<", form)

    def test_matrix_has_filters_and_gj_pill_bottom_right(self):
        html = self.client.get(self.url).content.decode("utf-8")
        seg = html.split("成本构成表", 1)[1]
        form = seg.split("</form>", 1)[0]
        self.assertIn('name="m_period"', form)
        self.assertIn('name="m_workshop"', form)
        self.assertIn(">筛选<", form)
        self.assertIn(">清空<", form)
        # 矩阵本体
        self.assertIn("total-cost-table", html)
        for subject in ("直接材料", "水电费", "折旧费", "人工费", "加工费"):
            self.assertIn(subject, seg.split("</table>", 1)[0], subject)
        # 勾稽胶囊在表尾右侧那块（pc-footer-status），不在标题行
        footer = seg.split('class="pc-footer"', 1)[1]
        status = footer.split('class="pc-footer-status"', 1)[1]
        self.assertIn("勾稽", status)
        heading = seg.split("</table>", 1)[0].split('class="section-heading"', 1)
        if len(heading) > 1:
            self.assertNotIn("勾稽", heading[1].split("</div>", 2)[0])

    def test_notes_are_short_text_not_tables(self):
        """三张表各有一段文字口径说明（用户 2026-09-24 加了「投料报工情形速查」）。"""
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertEqual(html.count("口径："), 3)
        self.assertEqual(html.count("pc-footer-notes"), 4)   # 3 处文案 + 1 条样式
        self.assertEqual(html.count('class="muted-text pc-footer-notes"'), 3)
        # 仍然没有「口径说明」四列表
        self.assertNotIn("口径说明", html)
        self.assertNotIn("basis-table", html)

    def test_issue_report_section_is_on_the_page(self):
        """用户 2026-09-24：生产成本页有「投料报工情形速查」，筛选字段＝期间/车间/工单号/情形。"""
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertIn("<h2>投料报工情形速查</h2>", html)
        # 标题下那句"每张工单一行的口径对照…"已按要求删除
        self.assertNotIn("这张单走了哪条核算路径", html)
        # 两张子表各有名字
        self.assertIn("情形小结", html)
        self.assertNotIn("情形小结（按序号）", html)
        self.assertIn("工单投料报工明细（可人工调整投料数量）", html)
        # 自己的一条筛选行：期间 / 车间 / 工单号 / 情形
        # （用户 2026-09-27：「含粗纱未报工单」勾选框已按你的要求删掉）
        forms = html.split('<form class="pc-filter-row"')[1:]
        self.assertEqual(len(forms), 3)
        issue_form = forms[2].split("</form>", 1)[0]
        for token in ('name="i_period"', 'name="i_workshop"', 'name="i_case"',
                      'name="i_order_no"',
                      ">筛选<", ">清空<", 'class="print-here"'):
            self.assertIn(token, issue_form, token)
        # 「含粗纱未报工单（N 张，不产生成本行）」这个字段删除
        self.assertNotIn("含粗纱未报工单", html)
        self.assertNotIn("不产生成本行", html)
        self.assertNotIn('name="i_roving"', issue_form)
        # 产品筛选已去掉
        self.assertNotIn('name="i_product"', issue_form)
        # 另外两张表的筛选条件原样带过去（三张表互不影响）
        for token in ('name="p_workshop"', 'name="m_workshop"'):
            self.assertIn(token, issue_form, token)

    def test_issue_report_detail_copies_work_order_cost_table(self):
        """用户 2026-09-24：明细表＝「工单成本明细」的同一套列（复制过来）+ 一列「情形」，
        只把「本期投料」做成可人工填写。

        用户 2026-09-27：原来的「期间」列改名为「核算期间」，并在它后面加一列「工单期间」。
        """
        html = self.client.get(self.url).content.decode("utf-8")
        detail = html.split("工单投料报工明细（可人工调整投料数量）", 1)[1]
        head = detail.split("<thead>", 1)[1].split("</thead>", 1)[0]
        for token in ("核算期间", "工单期间", "车间", "工单编号", "类别", "产品编码", "产品名称",
                      "计划完工时间", "实际完工时间", "完成状态", "计划数量",
                      "期初数量", "期初金额", "本期投料", "投料单价", "累计投料", "投料状态",
                      "本期完工", "累计完工", "本期完工率", "累计完工率", "是否用粗纱",
                      "报工数量", "在产品", "折合产量", "材料", "水电", "蒸汽", "天然气",
                      "折旧", "机物料", "染料助剂", "加工费", "工资",
                      "总成本", "完工成本", "单位成本", "标准单位成本", "单位成本差异",
                      "期末在产品数量", "在产品成本", "情形"):
            self.assertIn(">%s<" % token, head, token)
        # 只有「本期投料」这一列带输入框
        inputs = re.findall(r'name="(adj_[a-z_]+)_[^"]+"', detail)
        self.assertEqual(set(inputs), {"adj_issue_qty"})
        # 新表头顺序：核算期间 → 工单期间 → 车间 …
        heads = [re.sub(r"<[^>]+>", "", x).strip()
                 for x in re.findall(r"<th[^>]*>(.*?)</th>", head, re.S)]
        self.assertEqual(heads[:3], ["核算期间", "工单期间", "车间"])
        for order_no in ("MO-RanZheng", "MO-tefang"):
            self.assertIn('name="adj_issue_qty_%s"' % order_no, detail)
            self.assertIn('<input type="hidden" name="order_no" value="%s">' % order_no, detail)
        self.assertIn("保存人工调整并重算", detail)

    def test_issue_report_has_order_period_column(self):
        """「工单期间」＝工单明细上传的「创建时间」（ERP 建单时间）所属月份（用户 2026-09-27）。

        与「核算期间」（这张单的成本算在哪个月）区分开：一张 6 月下单、7 月发生成本的单，
        核算期间＝2026-07、工单期间＝2026-06。
        """
        from settlement.models import WorkOrderDetail

        batch = self.period.import_batches.first()
        WorkOrderDetail.objects.create(
            period=self.period, batch=batch, order_no="MO-tefang",
            created_at="2026-06-18", planned_finish="2026-06-30",
        )
        resp = self.client.get(self.url)
        by_order = {row.work_order.order_no: row for row in resp.context["detail_rows"]}
        row = by_order["MO-tefang"]
        self.assertEqual(row.detail_period_label, "2026-07")   # 核算期间
        self.assertEqual(row.order_period_label, "2026-06")    # 工单期间＝下单月份
        # 没有工单明细的单 → 空（页面显示 —）
        self.assertEqual(by_order["MO-RanZheng"].order_period_label, "")
        html = resp.content.decode("utf-8")
        self.assertIn('title="＝工单明细上传的「创建时间」', html)

    def test_order_period_label_parses_common_formats(self):
        from settlement.views import _order_period_label

        self.assertEqual(_order_period_label("2026-08-26"), "2026-08")
        self.assertEqual(_order_period_label("2026/8/26"), "2026-08")
        self.assertEqual(_order_period_label("20260826"), "2026-08")
        self.assertEqual(_order_period_label("2026年8月26日"), "2026-08")
        self.assertEqual(_order_period_label(""), "")
        self.assertEqual(_order_period_label(None), "")
        self.assertEqual(_order_period_label("不是日期"), "")

    def test_issue_report_manual_adjustment_saves_and_recalcs(self):
        """人工改投料/报工 → 存起来 + 该车间工单成本立即重算（用户 2026-09-24 方案 A）。"""
        from settlement.models import MaterialIssue, WorkOrderManualAdjustment, WorkReport

        order = WorkOrder.objects.get(period=self.period, order_no="MO-tefang")
        # 先给这张单真实的投料/报工台账（重算只认有业务的工单，否则结果会被清掉）
        MaterialIssue.objects.create(
            period=self.period, batch=self.period.import_batches.first(),
            work_order_no="MO-tefang", material_code="RM-T",
            material_name="原料T", actual_issue_quantity=Decimal("100"),
        )
        InventoryValuation.objects.create(
            period=self.period, material_code="RM-T",
            weighted_average_price=Decimal("10"),
        )
        WorkReport.objects.create(
            period=self.period, batch=self.period.import_batches.first(),
            work_order_no="MO-tefang", material_code="P-tefang", quantity=Decimal("80"),
            department="tefang",
        )

        # 人工把投料金额改成 1,234.56
        resp = self.client.post(
            self.url,
            {
                "order_no": "MO-tefang",
                "adj_issue_amount_MO-tefang": "1234.56",
            },
            follow=False,
        )
        self.assertEqual(resp.status_code, 302)      # PRG：保存后重定向
        adjustment = WorkOrderManualAdjustment.objects.get(
            period=self.period, order_no="MO-tefang"
        )
        self.assertEqual(adjustment.issue_amount, Decimal("1234.560000"))
        self.assertIsNone(adjustment.report_quantity)
        self.assertEqual(adjustment.updated_by, self.user)
        after = WorkOrderCostResult.objects.get(period=self.period, work_order=order)
        self.assertEqual(after.direct_material, Decimal("1234.56"))
        self.assertEqual(after.finished_quantity, Decimal("80.000000"))

        # 再人工把报工数量改成 500 → 完工数量跟着变
        self.client.post(
            self.url,
            {
                "order_no": "MO-tefang",
                "adj_issue_amount_MO-tefang": "1234.56",
                "adj_report_qty_MO-tefang": "500",
            },
        )
        after = WorkOrderCostResult.objects.get(period=self.period, work_order=order)
        self.assertEqual(after.finished_quantity, Decimal("500.000000"))

        # 清空 → 恢复系统值（调整记录删除 + 结果重算回去）
        self.client.post(
            self.url,
            {"order_no": "MO-tefang", "adj_report_qty_MO-tefang": ""},
        )
        self.assertFalse(
            WorkOrderManualAdjustment.objects.filter(
                period=self.period, order_no="MO-tefang"
            ).exists()
        )
        restored = WorkOrderCostResult.objects.get(period=self.period, work_order=order)
        self.assertEqual(restored.finished_quantity, Decimal("80.000000"))
        # 人工值不再生效（回到按账单重算的口径）
        self.assertNotEqual(restored.direct_material, Decimal("1234.56"))

    def test_issue_report_only_fills_issue_quantity_needs_a_price(self):
        """只填投料数量：该单原无投料、所在车间本月也没投料 → 取不到单价 → 报错、不保存。"""
        from settlement.models import WorkOrderManualAdjustment

        resp = self.client.post(
            self.url,
            {"order_no": "MO-RanZheng", "adj_issue_qty_MO-RanZheng": "50"},
            follow=True,
        )
        self.assertEqual(resp.status_code, 200)
        self.assertContains(resp, "取不到投料单价")
        self.assertFalse(
            WorkOrderManualAdjustment.objects.filter(
                period=self.period, order_no="MO-RanZheng"
            ).exists()
        )

    def test_issue_report_only_fills_issue_quantity_uses_workshop_average(self):
        """只填投料数量，该单原无投料、但同车间别的工单有投料 → 按车间均价折算金额。"""
        from settlement.models import MaterialIssue, WorkOrderManualAdjustment

        # 同车间（染色）另一张工单有投料：100 kg × 10 元 → 车间均价 10 元/kg
        other = WorkOrder.objects.create(
            period=self.period, batch=self.period.import_batches.first(),
            order_no="MO-RZ-2", production_dept_code="RanZheng",
            plan_quantity=Decimal("10"),
        )
        MaterialIssue.objects.create(
            period=self.period, batch=self.period.import_batches.first(),
            work_order_no=other.order_no, material_code="RM-T",
            material_name="原料T", actual_issue_quantity=Decimal("100"),
        )
        InventoryValuation.objects.create(
            period=self.period, material_code="RM-T",
            weighted_average_price=Decimal("10"),
        )
        resp = self.client.post(
            self.url,
            {"order_no": "MO-RanZheng", "adj_issue_qty_MO-RanZheng": "50"},
            follow=False,
        )
        self.assertEqual(resp.status_code, 302)
        adjustment = WorkOrderManualAdjustment.objects.get(
            period=self.period, order_no="MO-RanZheng"
        )
        self.assertEqual(adjustment.issue_quantity, Decimal("50.000000"))
        # 金额＝50 × 车间均价 10
        self.assertEqual(adjustment.issue_amount, Decimal("500.000000"))

    def test_issue_report_detail_matches_work_order_cost_page(self):
        """明细表直接复用「工单成本明细」的取数：同一张工单两边数字一致。"""
        resp = self.client.get(self.url, {"i_workshop": "tefang"})
        detail = {row.work_order.order_no: row for row in resp.context["detail_rows"]}
        self.assertIn("MO-tefang", detail)
        woc = self.client.get(
            reverse("work_order_costs", args=[self.period.id]),
            {"w_workshop": "tefang"},
        )
        page_rows = {r.work_order.order_no: r for r in woc.context["rows"]}
        mine, theirs = detail["MO-tefang"], page_rows["MO-tefang"]
        for field in ("direct_material", "finished_cost", "wip_cost", "total_cost",
                      "finished_quantity", "wip_quantity", "current_issue",
                      "opening_quantity", "opening_amount", "accum_issue",
                      "current_unit_price", "equivalent_qty", "rate_wages_pct"):
            self.assertEqual(
                Decimal(str(getattr(mine, field) or 0)),
                Decimal(str(getattr(theirs, field) or 0)),
                field,
            )
        self.assertEqual(mine.wo_category, theirs.wo_category)
        self.assertEqual(mine.uses_roving, theirs.uses_roving)
        self.assertEqual(mine.case_label, CASE_NORMAL)

    def test_issue_report_manual_flag_shows_on_page(self):
        """人工调整过的行会打标记，并在提示里写明以人工值为准。"""
        from settlement.models import WorkOrderManualAdjustment

        WorkOrderManualAdjustment.objects.create(
            period=self.period, order_no="MO-tefang", issue_amount=Decimal("888"),
        )
        resp = self.client.get(self.url, {"i_order_no": "MO-tefang"})
        rows = resp.context["issue_rows"]
        self.assertEqual(len(rows), 1)
        self.assertTrue(rows[0]["has_manual"])
        self.assertEqual(rows[0]["manual_issue_amount"], Decimal("888"))
        self.assertEqual(rows[0]["issue_amount"], Decimal("888"))
        self.assertIn("已人工调整", rows[0]["flags"])

    def test_issue_report_rows_and_case_filter(self):
        """速查表的逐单分类与情形筛选（数据来自 build_issue_report_rows）。"""
        resp = self.client.get(self.url)
        rows = resp.context["issue_rows"]
        # setUp 造了 2 张（染色/纺纱），有完工但无投料、也无报工 → 不属任何要盯的情形
        self.assertEqual(len(rows), 2)
        self.assertEqual({r["case"] for r in rows}, {CASE_NORMAL})
        self.assertEqual({r["combo"] for r in rows}, {"无投料+无报工"})
        # 情形小结：4 类（本期没有的也列出来，工单数 0）+ 一行「⑤ 正常投料生产工单」，
        # 合计＝明细行数（用户 2026-09-24：合计要和列出来的行对得上）
        summary = resp.context["issue_case_summary"]
        # 前 5 行固定是 ①~⑤；若本期有 BOM 缺单价的单，最后再多一行「⑥ BOM 子项缺单价的工单」
        self.assertGreaterEqual(len(summary), len(CASE_CHOICES) + 1)
        self.assertEqual(
            [item["case"] for item in summary[: len(CASE_CHOICES) + 1]],
            list(CASE_CHOICES) + [CASE_OTHER],
        )
        self.assertEqual({item["count"] for item in summary[: len(CASE_CHOICES)]}, {0})
        self.assertEqual(summary[len(CASE_CHOICES)]["count"], 2)   # 2 张都归"⑤ 正常投料生产工单"
        # 合计＝明细行数（缺价那一行是标注类，不计入）
        self.assertEqual(
            sum(item["count"] for item in summary if not item.get("overlap")), 2
        )
        # 「在产品 / 产成品成本怎么算」要点（用户 2026-09-30）：每种情形都要有，且写得对得上结算口径
        from settlement.services.workorder_costing import (
            CASE_COST_BASIS, CASE_REPORT_EXCEED, CASE_REPORT_ONLY,
        )

        for item in summary:
            self.assertTrue(item["basis"], f"{item['case']} 缺少成本口径要点")
        self.assertEqual(
            [item["basis"] for item in summary[: len(CASE_CHOICES) + 1]],
            [CASE_COST_BASIS[case] for case in list(CASE_CHOICES) + [CASE_OTHER]],
        )
        self.assertIn("全部算产成品", CASE_COST_BASIS[CASE_REPORT_ONLY])     # ④ 反推料全进产成品
        self.assertIn("在产品＝0", CASE_COST_BASIS[CASE_REPORT_ONLY])        # ④ 没有在产品（用户 2026-09-30）
        self.assertIn("在产品＝0", CASE_COST_BASIS[CASE_REPORT_EXCEED])      # ③ 也不确认在产品
        self.assertIn("补", CASE_SUGGESTIONS[CASE_REPORT_EXCEED])            # ③ 建议里说明已自动补料
        self.assertIn("折合产量", CASE_COST_BASIS[CASE_OTHER])
        # 在产品 / 产成品口径写在「建议处理方法」那一列里（不另开列）
        html = resp.content.decode("utf-8")
        self.assertIn("建议处理方法（含在产品 / 产成品成本口径）", html)
        self.assertNotIn("在产品 / 产成品成本怎么算</th>", html)
        self.assertIn("剩余纱线", html)
        # 需要人工介入的数量＝除"其他"以外的合计
        self.assertEqual(resp.context["issue_attention_count"], 0)
        # 按情形筛选：命中/不命中
        resp = self.client.get(self.url, {"i_case": CASE_NORMAL})
        self.assertEqual(len(resp.context["issue_rows"]), 2)
        resp = self.client.get(self.url, {"i_case": CASE_REPORT_ONLY})
        self.assertEqual(len(resp.context["issue_rows"]), 0)
        # 按车间 / 工单号筛选
        resp = self.client.get(self.url, {"i_workshop": "tefang"})
        self.assertEqual(len(resp.context["issue_rows"]), 1)
        resp = self.client.get(self.url, {"i_order_no": "MO-tefang"})
        rows = resp.context["issue_rows"]
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["order_no"], "MO-tefang")
        self.assertEqual(rows[0]["plan_qty"], Decimal("100.00"))

    def test_issue_report_only_counts_orders_with_cost(self):
        """统计口径（用户 2026-09-24）：默认只列"当月发生核算成本的工单"。

        BC- 粗纱工单本期没报工 → 计算层按计划自动完工、不产生成本行 → 默认不列入；
        勾选「含粗纱未报工单」后才补出来。
        """
        batch = self.period.import_batches.first()
        roving = WorkOrder.objects.create(
            period=self.period, batch=batch, order_no="BC-16.9999CS-9900001",
            production_dept_code="tefang", plan_quantity=Decimal("1000"),
        )
        BomLine.objects.create(
            period=self.period, batch=batch, work_order_no=roving.order_no,
            material_code="BC-16.9999CS", sub_item_code="RM-ROV",
            qty_numerator=Decimal("1"), qty_denominator=Decimal("1"),
        )
        # 默认：不含（它没有成本结果）
        resp = self.client.get(self.url)
        self.assertEqual(resp.context["issue_totals_display"]["roving_only_count"], 1)
        self.assertEqual(len(resp.context["issue_rows"]), 2)
        self.assertNotIn(
            roving.order_no,
            [row["order_no"] for row in resp.context["issue_rows"]],
        )
        # 勾选后才列出
        resp = self.client.get(self.url, {"i_roving": "1"})
        self.assertEqual(len(resp.context["issue_rows"]), 3)
        self.assertIn(
            roving.order_no,
            [row["order_no"] for row in resp.context["issue_rows"]],
        )

    def test_issue_report_classifies_issue_without_report(self):
        """有投料、无报工：当月有在产属正常 → 归"正常"，只有计划完工已过 3 个月才进①。

        材料成本＝实领金额（全额留在产品，完工=0）。
        """
        from settlement.models import MaterialIssue

        order = WorkOrder.objects.get(period=self.period, order_no="MO-RanZheng")
        MaterialIssue.objects.create(
            period=self.period, batch=self.period.import_batches.first(),
            work_order_no=order.order_no, material_code="RM-X",
            material_name="原料X", actual_issue_quantity=Decimal("100"),
        )
        InventoryValuation.objects.create(
            period=self.period, material_code="RM-X",
            weighted_average_price=Decimal("10"),
        )
        resp = self.client.get(self.url, {"i_order_no": order.order_no})
        rows = resp.context["issue_rows"]
        self.assertEqual(len(rows), 1)
        row = rows[0]
        self.assertEqual(row["combo"], "有投料+无报工")
        # 该工单计划完工时间在核算月内（未超 3 个月）→ 不算异常
        self.assertEqual(row["case"], CASE_NORMAL)
        self.assertEqual(row["issue_amount"], Decimal("1000.00"))
        # 材料来源归因：全部是实领
        self.assertEqual(row["from_issue"], Decimal("1000.00"))
        self.assertEqual(row["from_calc"], Decimal("0"))
        # 完工＝0（没有报工）；表里的成本三列直接取结算结果
        self.assertEqual(row["report_qty"], Decimal("0"))
        self.assertEqual(row["direct_material"], Decimal("1000.00"))

    def test_report_only_requires_never_issued(self):
        """「无投料-有报工」只收**从来没有投过料**的单（用户 2026-09-27）。

        以前投过料（投料进度表／工单明细的投料数量／期初在产品里有材料）的单，
        本期虽然没投料、有报工，也**不算**这一类，归「— 正常」并打提示。
        """
        from settlement.models import IssueProgress, WorkReport

        batch = self.period.import_batches.first()
        # ① 纯无投料：只有报工，任何地方都没有投料记录
        pure = WorkOrder.objects.get(period=self.period, order_no="MO-tefang")
        WorkReport.objects.create(
            period=self.period, batch=batch, work_order_no=pure.order_no,
            material_code="P-tefang", quantity=Decimal("10"), department="tefang",
        )
        # ② 以前投过料（投料进度表里有）→ 不算「无投料」
        prior = WorkOrder.objects.get(period=self.period, order_no="MO-RanZheng")
        WorkReport.objects.create(
            period=self.period, batch=batch, work_order_no=prior.order_no,
            material_code="P-RanZheng", quantity=Decimal("10"), department="RanZheng",
        )
        IssueProgress.objects.create(
            period=self.period, batch=batch, work_order_no=prior.order_no,
            material_code="RM-P", actual_issue_quantity=Decimal("50"),
        )
        resp = self.client.get(self.url)
        rows = {row["order_no"]: row for row in resp.context["issue_rows"]}
        self.assertEqual(rows[pure.order_no]["case"], CASE_REPORT_ONLY)
        self.assertEqual(rows[pure.order_no]["combo"], "无投料+有报工")
        self.assertEqual(rows[prior.order_no]["case"], CASE_NORMAL)
        # 行里的 flags 是「；」拼好的字符串
        self.assertIn("以前投过料", str(rows[prior.order_no]["flags"]))

    def test_case_two_is_long_pending_wip(self):
        """② 改口径（用户 2026-09-27）：在产品 > 0 且计划完工已过 3 个月 → ②，
        **不再要求本月有投料**；不到 3 个月的仍归「正常投料生产工单」。"""
        # setUp 的 MO-RanZheng：计划完工在 2026-02-28（已过 3 个月），有在产品 20
        order = WorkOrder.objects.get(period=self.period, order_no="MO-RanZheng")
        order.planned_finish = timezone.make_aware(datetime(2026, 2, 28, 8, 0))
        order.save(update_fields=["planned_finish"])
        resp = self.client.get(self.url)
        rows = {row["order_no"]: row for row in resp.context["issue_rows"]}
        self.assertEqual(rows["MO-RanZheng"]["case"], CASE_NO_REPORT)
        # 本月没有投料、也没有报工，靠"在产品还挂着"进 ②
        self.assertEqual(rows["MO-RanZheng"]["issue_qty"], Decimal("0"))
        self.assertEqual(rows["MO-RanZheng"]["report_qty"], Decimal("0"))
        self.assertGreater(rows["MO-RanZheng"]["wip_qty"], Decimal("0"))
        self.assertIn("在产品仍挂着", str(rows["MO-RanZheng"]["flags"]))
        # 计入了「需要人工介入」
        self.assertGreaterEqual(resp.context["issue_attention_count"], 1)

    def test_case_two_needs_wip(self):
        """② 的两种触发都不成立时归「正常」：在产品 = 0 且本月既没投料也没报工。"""
        order = WorkOrder.objects.get(period=self.period, order_no="MO-tefang")
        order.planned_finish = timezone.make_aware(datetime(2025, 1, 31, 8, 0))
        order.save(update_fields=["planned_finish"])
        result = WorkOrderCostResult.objects.get(
            period=self.period, work_order=order
        )
        result.wip_quantity = Decimal("0")
        result.save(update_fields=["wip_quantity"])
        resp = self.client.get(self.url)
        rows = {row["order_no"]: row for row in resp.context["issue_rows"]}
        self.assertEqual(rows["MO-tefang"]["case"], CASE_NORMAL)

    def test_case_two_also_covers_former_case_five(self):
        """原⑤「有投料-部分报工」已并入 ②（用户 2026-09-27）：
        本月投料了、也报了工、但报工 < 计划，且计划完工已过 3 个月 → 归 ②（哪怕在产品=0）。"""
        from settlement.models import MaterialIssue, WorkReport

        batch = self.period.import_batches.first()
        order = WorkOrder.objects.get(period=self.period, order_no="MO-tefang")
        order.planned_finish = timezone.make_aware(datetime(2026, 2, 28, 8, 0))
        order.save(update_fields=["planned_finish"])
        # 在产品设为 0：证明命中的是"本月投料但没报完"这一条，而不是"在产品挂着"
        result = WorkOrderCostResult.objects.get(period=self.period, work_order=order)
        result.wip_quantity = Decimal("0")
        result.save(update_fields=["wip_quantity"])
        MaterialIssue.objects.create(
            period=self.period, batch=batch, work_order_no=order.order_no,
            material_code="RM-F", material_name="原料F",
            actual_issue_quantity=Decimal("100"),
        )
        InventoryValuation.objects.create(
            period=self.period, material_code="RM-F",
            weighted_average_price=Decimal("10"),
        )
        WorkReport.objects.create(
            period=self.period, batch=batch, work_order_no=order.order_no,
            material_code="P-tefang", quantity=Decimal("80"), department="tefang",
        )
        resp = self.client.get(self.url)
        row = {r["order_no"]: r for r in resp.context["issue_rows"]}[order.order_no]
        self.assertEqual(row["case"], CASE_NO_REPORT)
        # 报工(80) < 计划(100)，且没有"报工超投料量" → 不是 ③
        self.assertLess(row["report_qty"], row["plan_qty"])
        self.assertEqual(row["issue_qty"], Decimal("100"))
        self.assertIn("未报完", str(row["flags"]))
    def test_case_advice_texts_and_other_label(self):
        """② 的名称/建议按用户 2026-09-27 定稿（原⑤已并入 ②，情形 5 类 → 4 类，
        「正常投料生产工单」加序号 ⑤）。"""
        self.assertEqual(CASE_OTHER, "⑤ 正常投料生产工单")
        # ② 名称＝「有投料-无报工」，括号里的文案保持上一轮定下的不变
        self.assertEqual(
            CASE_NO_REPORT, "② 有投料-无报工（在产长期挂账超计划完工3个月以上）"
        )
        # 情形只剩 4 类（原⑤已并入 ②）
        self.assertEqual(len(CASE_CHOICES), 4)
        for choice in CASE_CHOICES:
            self.assertFalse(choice.startswith("⑤"))
        # 原⑤的编号 ⑤ 已让位给「口径说明」里的第 5 条，所以不再作为情形出现
        advice_no_report = CASE_SUGGESTIONS[CASE_NO_REPORT]
        self.assertIn("在产挂了 3 个月以上", advice_no_report)
        self.assertIn("人工核查", advice_no_report)
        self.assertIn("工单投料", advice_no_report)
        # 原⑤的情况也被这条建议覆盖（本月还在投料但没报完）
        self.assertIn("本月还在投料", advice_no_report)
        # 用户 2026-09-27：这两句已删除
        self.assertNotIn("计划完工时间是旧的", advice_no_report)
        self.assertNotIn("按完工入库结转", advice_no_report)
        self.assertNotIn("别再挂着", advice_no_report)
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertIn("正常投料生产工单", html)
        self.assertNotIn("其他（无需处理）", html)
        self.assertNotIn("先查是不是漏报工", html)
        # 页面上不再有 ⑤ 类情形
        self.assertNotIn("⑤ 有投料-部分报工", html)
        self.assertIn("情形共 4 类", html)

    def test_issue_unit_price_keeps_two_decimals(self):
        """投料单价保留 2 位小数（用户 2026-09-27），页面上就是 2 位。"""
        from settlement.models import InventoryValuation, MaterialIssue

        batch = self.period.import_batches.first()
        order = WorkOrder.objects.get(period=self.period, order_no="MO-tefang")
        # 两种料各投 1：加权平均 = (100 + 33.333333)/2 = 66.6667 → 保留 2 位 66.67
        for code, price in (("RM-P", Decimal("100")),
                            ("RM-Q", Decimal("33.333333"))):
            MaterialIssue.objects.create(
                period=self.period, batch=batch, work_order_no=order.order_no,
                material_code=code, material_name=code, actual_issue_quantity=Decimal("1"),
            )
            InventoryValuation.objects.create(
                period=self.period, material_code=code, weighted_average_price=price,
            )
        resp = self.client.get(self.url)
        row = {r.work_order.order_no: r for r in resp.context["detail_rows"]}[order.order_no]
        self.assertEqual(row.current_unit_price, Decimal("66.67"))
        self.assertFalse(row.current_unit_price_from_bom)
        html = resp.content.decode("utf-8")
        self.assertIn("66.67", html)
        self.assertNotIn("66.6667", html)

    def test_issue_unit_price_falls_back_to_bom(self):
        """本期没有投料的单，「投料单价」按 BOM 表折算（用户 2026-09-27），不再显示 —。"""
        from settlement.models import BomLine, InventoryValuation

        batch = self.period.import_batches.first()
        order = WorkOrder.objects.get(period=self.period, order_no="MO-tefang")
        for code, qty, price in (("RM-B1", Decimal("0.5"), Decimal("100")),
                                 ("RM-B2", Decimal("0.5"), Decimal("20"))):
            BomLine.objects.create(
                period=self.period, batch=batch, work_order_no=order.order_no,
                material_code="P-X", sub_item_code=code,
                qty_numerator=qty, qty_denominator=Decimal("1"),
            )
            InventoryValuation.objects.create(
                period=self.period, material_code=code,
                weighted_average_price=price,
            )
        resp = self.client.get(self.url)
        row = {r.work_order.order_no: r for r in resp.context["detail_rows"]}[order.order_no]
        # Σ(用量×价)=0.5×100+0.5×20=60；Σ用量=1 → 60.00
        self.assertEqual(row.current_unit_price, Decimal("60.00"))
        self.assertTrue(row.current_unit_price_from_bom)
        html = resp.content.decode("utf-8")
        self.assertIn("按 BOM 表折算", html)   # 悬停提示

    def test_issue_unit_price_export_is_two_decimals(self):
        """导出里的「投料单价」也是 2 位小数（保留 2 位、取不到价显示 —）。"""
        from settlement.models import BomLine, InventoryValuation

        batch = self.period.import_batches.first()
        order = WorkOrder.objects.get(period=self.period, order_no="MO-tefang")
        # 两个子项各 1 用量：(100 + 33.333333)/2 = 66.6667 → 保留 2 位 66.67
        for code, price in (("RM-E1", Decimal("100")),
                            ("RM-E2", Decimal("33.333333"))):
            BomLine.objects.create(
                period=self.period, batch=batch, work_order_no=order.order_no,
                material_code="P-X", sub_item_code=code,
                qty_numerator=Decimal("1"), qty_denominator=Decimal("1"),
            )
            InventoryValuation.objects.create(
                period=self.period, material_code=code, weighted_average_price=price,
            )
        resp = self.client.get(reverse("issue_report_export", args=[self.period.id]))
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(BytesIO(resp.content))
        sheet = wb["工单投料报工明细"]
        headers = [c.value for c in sheet[1]]
        col = headers.index("投料单价") + 1
        values = {}
        for row in sheet.iter_rows(min_row=2, values_only=True):
            values[row[3]] = row[col - 1]
        # (100 + 33.333333)/2 = 66.6667 → 保留 2 位 66.67
        self.assertEqual(Decimal(str(values["MO-tefang"])), Decimal("66.67"))
        # 另一张单没有 BOM 也没有投料 → 显示 —
        self.assertEqual(values["MO-RanZheng"], "—")

    def test_normal_row_got_a_serial_number(self):
        """「正常投料生产工单」也要有序号（用户 2026-09-27）：接在 ①~④ 后面是 ⑤；

        2026-09-30 又加了第 6 种「⑥ BOM 子项缺单价的工单」，排在 ⑤ 后面。
        """
        resp = self.client.get(self.url)
        summary = resp.context["issue_case_summary"]
        cases = [item["case"] for item in summary]
        # 小结按 ①②③④ + ⑤ 正常 + ⑥ 缺单价 排列
        self.assertEqual(cases[:4], list(CASE_CHOICES))
        self.assertEqual(cases[4], "⑤ 正常投料生产工单")
        self.assertEqual(cases[5], CASE_NO_PRICE)
        # 情形×车间表的最后两行也是带序号的
        matrix_cases = [row["case"] for row in resp.context["case_workshop"]["rows"]]
        self.assertEqual(matrix_cases[-2:], ["⑤ 正常投料生产工单", CASE_NO_PRICE])
        # 明细表「情形」列的取值不受影响（还是「— 正常」）
        self.assertIn(CASE_NORMAL, [r["case"] for r in resp.context["issue_rows"]])
        html = resp.content.decode("utf-8")
        self.assertIn("⑤ 正常投料生产工单", html)
        # 口径说明里也点到 ⑤
        self.assertIn("⑤ 正常投料生产工单</b>就是本月正常的单", html)

    def test_issue_report_export_is_single_sheet(self):
        """「二、工单投料报工明细」自己的导出（用户 2026-09-27）：
        一张工作表、列与页面一致（含核算期间/工单期间/情形），末行合计。"""
        resp = self.client.get(reverse("issue_report_export", args=[self.period.id]))
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(BytesIO(resp.content))
        self.assertEqual(wb.sheetnames, ["工单投料报工明细"])
        sheet = wb["工单投料报工明细"]
        headers = [cell.value for cell in sheet[1]]
        self.assertEqual(headers[:3], ["核算期间", "工单期间", "车间"])
        for token in ("本期投料", "累计投料", "完工成本", "在产品成本", "情形"):
            self.assertIn(token, headers)
        # 表头 1 行 + 2 张工单 + 合计 1 行
        self.assertEqual(sheet.max_row, 4)
        self.assertTrue(str(sheet.cell(row=4, column=1).value).startswith("合计"))

    def test_matrix_export_is_single_sheet(self):
        """「成本构成表」自己的导出（用户 2026-09-27）：只有一张工作表。"""
        resp = self.client.get(
            reverse("production_cost_matrix_export", args=[self.period.id])
        )
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(BytesIO(resp.content))
        self.assertEqual(wb.sheetnames, ["成本构成表"])
        self.assertEqual(wb["成本构成表"]["A3"].value, "成本项目")

    def test_page_has_section_export_buttons_and_short_heading(self):
        """标题改名；「成本构成表」和「二、工单投料报工明细」各有一个「导出Excel」。"""
        html = self.client.get(self.url).content.decode("utf-8")
        # 用户 2026-09-27：原名「成本项目 × 四大块」太土 → 改名「成本构成表」
        self.assertIn("<h2>成本构成表</h2>", html)
        self.assertNotIn("成本项目 × 四大块", html)
        self.assertNotIn("四大块", html)
        self.assertNotIn("期初 / 本期投入 / 产成品 / 在产品", html)
        self.assertIn(
            reverse("production_cost_matrix_export", args=[self.period.id]), html
        )
        self.assertIn(reverse("issue_report_export", args=[self.period.id]), html)
        # 页头一个 + 两张表各一个
        self.assertGreaterEqual(html.count(">导出Excel<"), 3)

    def test_case_workshop_matrix_breaks_down_by_workshop(self):
        """情形 × 车间（工单数）（用户 2026-09-27）：
        能看到每类情形分别在哪些车间、各多少张；合计＝明细行数。"""
        resp = self.client.get(self.url)
        matrix = resp.context["case_workshop"]
        # setUp 有两张单（染色/纺纱），都归「正常投料生产工单」
        labels = [item["workshop_label"] for item in matrix["workshops"]]
        self.assertIn("染色车间", labels)
        self.assertIn("纺纱车间", labels)
        normal_row = next(
            row for row in matrix["rows"] if row["case"] == CASE_OTHER
        )
        by_ws = {cell["workshop"]: cell["count"] for cell in normal_row["cells"]}
        self.assertEqual(by_ws.get("RanZheng"), 1)
        self.assertEqual(by_ws.get("tefang"), 1)
        self.assertEqual(normal_row["count"], 2)
        # 合计＝各车间之和＝明细行数
        self.assertEqual(matrix["total"], 2)
        self.assertEqual(
            sum(item["count"] for item in matrix["totals"]), 2
        )

    def test_case_workshop_matrix_follows_workshop_filter(self):
        """情形×车间 跟着「车间」筛选走：筛纺纱后只剩纺纱那一列。"""
        resp = self.client.get(self.url, {"i_workshop": "tefang"})
        matrix = resp.context["case_workshop"]
        self.assertEqual(
            [item["workshop"] for item in matrix["workshops"]], ["tefang"]
        )
        self.assertEqual(matrix["total"], 1)

    def test_empty_case_advice_is_not_greyed(self):
        """本期没有的情形（工单数 0）：建议处理方法照常显示，不要整行发灰（用户 2026-09-27）。"""
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertIn(
            ".pc-case-table tr.pc-case-empty td.pc-advice{color:#3f4d5a;}", html
        )

    def test_save_hint_sits_next_to_the_button(self):
        """「清空＝恢复系统值…」这段提示要贴在保存按钮旁边（用户 2026-09-27）。

        原来是 .ir-hint{flex:1 1 420px} + 空的「有 N 个格子改过」占位，
        会把这段说明顶到整行最右边；现在空提示不占位、说明紧跟按钮。
        """
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertIn('id="ir-save-note"', html)
        self.assertIn(".ir-actions{display:flex;flex-wrap:wrap;gap:8px 14px", html)
        self.assertIn("#ir-dirty-hint{flex:0 0 auto;}", html)
        # 保存按钮后面紧跟提示（中间只隔一个空提示 span）
        seg = html.split("保存人工调整并重算", 1)[1]
        self.assertIn('id="ir-save-note"', seg.split("</div>", 1)[0])
        self.assertIn("清空＝恢复系统值", seg.split("</div>", 1)[0])

    def test_helper_sentences_are_removed(self):
        """用户 2026-09-27 要求删掉两句提示说明（小结上方那句、车间分布表上方那句）。"""
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertNotIn("默认显示", html)
        self.assertNotIn("下表跟着下方「核算期间 / 车间」筛选走", html)
        self.assertNotIn("各情形分别在哪个车间、各多少张", html)
        self.assertNotIn("点上面的情形名可以把下方明细筛成这一类", html)
        # 车间分布表本身还在（只删了那句说明）
        self.assertIn("工单数", html)

    def test_issue_report_filter_labels_say_accounting_period(self):
        """明细表的期间筛选改叫「核算期间」，默认本期但不写"默认"两字（用户 2026-09-27）。"""
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertIn("核算期间", html)
        self.assertIn("本期（2026-07）", html)
        self.assertNotIn("本期（2026-07，默认）", html)

    def test_order_period_filter_narrows_the_detail_table(self):
        """明细表新增「工单期间」筛选字段（用户 2026-09-27）：
        选了某个下单月份，明细只剩那个月下的单；情形小结不受影响。"""
        from settlement.models import WorkOrderDetail

        batch = self.period.import_batches.first()
        WorkOrderDetail.objects.create(
            period=self.period, batch=batch, order_no="MO-tefang",
            created_at="2026-04-11",
        )
        resp = self.client.get(self.url)
        html = resp.content.decode("utf-8")
        self.assertIn('name="i_order_period"', html)
        self.assertIn("工单期间", html)
        # 下拉只列**本期这张表里真的出现过**的下单月份：本测试只给 MO-tefang 挂了 2026-04
        self.assertEqual(resp.context["order_period_options"], ["2026-04"])
        # 不筛：两张都在
        self.assertEqual(len(resp.context["detail_rows"]), 2)
        # 筛 2026-04：只剩 MO-tefang 那张
        resp = self.client.get(self.url, {"i_order_period": "2026-04"})
        self.assertEqual(
            [row.work_order.order_no for row in resp.context["detail_rows"]],
            ["MO-tefang"],
        )
        self.assertEqual(resp.context["detail_totals"]["count"], 1)
        # 小结（按核算期间）不受工单期间筛选影响
        self.assertEqual(resp.context["issue_totals"]["count"], 2)

    def test_case_link_jumps_to_detail_table_not_page_top(self):
        """点情形名要跳到「二、工单投料报工明细」，不能回到页顶（用户 2026-09-27）。

        做法：明细小节挂 id="issue-detail"，情形链接带 #issue-detail；
        并顺手清掉 base.html 的滚动位置缓存，免得它把锚点跳转覆盖掉。
        """
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertIn('id="issue-detail"', html)
        self.assertIn("#issue-detail", html)
        # 情形链接（小结那一列）必须带锚点
        link = re.search(
            r'<a class="document-link"[^>]*href="([^"]*#issue-detail)"', html
        )
        self.assertIsNotNone(link, "情形链接没有带 #issue-detail")
        self.assertIn("i_case=", link.group(1))
        # 清滚动缓存，避免锚点被「恢复上次滚动位置」覆盖
        self.assertIn("sessionStorage.removeItem('keepScrollPos')", html)
        # 明细小节刚好在 id 上（标题就是明细表那一行）
        self.assertRegex(
            html, r'<h3 class="pc-subheading" id="issue-detail"[^>]*>二、工单投料报工明细'
        )

    def test_filtered_detail_shows_which_case_is_active(self):
        """按情形筛过来后，明细上方要明确显示「已按情形筛选：X」并给一键恢复。"""
        resp = self.client.get(self.url, {"i_case": CASE_NORMAL})
        html = resp.content.decode("utf-8")
        self.assertIn("已按情形筛选", html)
        self.assertIn(CASE_NORMAL, html)
        self.assertIn("只看全部情形", html)
        # 没筛情形时不出现这条提示
        plain = self.client.get(self.url).content.decode("utf-8")
        self.assertNotIn("已按情形筛选", plain)

    def test_issue_report_long_pending_goes_into_case_one(self):
        """有投料-无报工 + 计划完工已过 3 个月以上 → ①（长期挂账要盯）。"""
        from settlement.models import MaterialIssue

        order = WorkOrder.objects.get(period=self.period, order_no="MO-RanZheng")
        order.planned_finish = timezone.make_aware(
            datetime(2026, 2, 28, 8, 0)
        )
        order.save(update_fields=["planned_finish"])
        MaterialIssue.objects.create(
            period=self.period, batch=self.period.import_batches.first(),
            work_order_no=order.order_no, material_code="RM-X",
            material_name="原料X", actual_issue_quantity=Decimal("100"),
        )
        InventoryValuation.objects.create(
            period=self.period, material_code="RM-X",
            weighted_average_price=Decimal("10"),
        )
        resp = self.client.get(self.url, {"i_order_no": order.order_no})
        rows = resp.context["issue_rows"]
        self.assertEqual(rows[0]["case"], CASE_NO_REPORT)
        self.assertIn("计划完工已过", rows[0]["flags"])
        # 小结里 ① 有 1 张
        summary = {item["case"]: item["count"] for item in resp.context["issue_case_summary"]}
        self.assertEqual(summary[CASE_NO_REPORT], 1)

    def test_workshop_and_period_filters_scope_each_table(self):
        # 车间汇总：只看纺纱车间（行 + 合计都只剩它）
        resp = self.client.get(self.url, {"p_workshop": "tefang"})
        rows = resp.context["rows"]
        self.assertEqual([row["workshop"] for row in rows], ["tefang"])
        self.assertEqual(resp.context["totals"]["order_count"], 1)
        # 矩阵：只看纺纱车间（表头只剩纺纱 + 合计）
        resp2 = self.client.get(self.url, {"m_workshop": "tefang"})
        html2 = resp2.content.decode("utf-8")
        head = html2.split(
            'class="data-table total-cost-table"', 1
        )[1].split("</thead>", 1)[0]
        self.assertIn("纺纱车间", head)
        self.assertNotIn("染色车间", head)
        # 两张表互不影响：矩阵筛选不改变车间汇总的行
        self.assertEqual(len(resp2.context["rows"]), 2)
        # 期间筛选：建一个月度数据后，选它能看到那个月的数据（默认仍是本期）
        june = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        june_batch = self.create_batch(
            june, ImportBatch.Kind.WORK_ORDER, "pc-merge-june"
        )
        june_order = WorkOrder.objects.create(
            period=june, batch=june_batch, order_no="MO-JUNE",
            production_dept_code="BaoFu", plan_quantity=Decimal("10"),
        )
        WorkOrderCostResult.objects.create(
            period=june, work_order=june_order, product_code="P-JUNE",
            product_name="六月产品", plan_quantity=Decimal("10"),
            finished_quantity=Decimal("10"), wip_quantity=Decimal("0"),
            direct_material=Decimal("100"), finished_material=Decimal("100"),
            finished_cost=Decimal("100"), wip_cost=Decimal("0"),
            total_cost=Decimal("100"),
        )
        resp3 = self.client.get(self.url, {"p_period": "2026-06"})
        self.assertEqual(resp3.context["summary_period_label"], "2026-06")
        self.assertEqual(
            [row["workshop"] for row in resp3.context["rows"]], ["BaoFu"]
        )
        # 期间下拉里能看到 2026-06；默认（不传参）仍只显示本期
        self.assertIn(
            "2026-06",
            [item["value"] for item in resp3.context["period_options"]],
        )
        self.assertEqual(
            {row["workshop"] for row in self.client.get(self.url).context["rows"]},
            {"RanZheng", "tefang"},
        )

    def test_workshop_rows_follow_canonical_order(self):
        """车间汇总的行顺序＝染色 → 纺纱 → 包覆 → 针织 → 机织 → 浆粕（用户 2026-09-21）。"""
        batch = self.create_batch(self.period, ImportBatch.Kind.WORK_ORDER, "pc-order")
        # 补齐其余车间：车间要出现在汇总里，必须有「工单结算结果」或「车间费用池」
        for code in ("BaoFu", "ZhenZhi", "JiZhi"):
            order = WorkOrder.objects.create(
                period=self.period, batch=batch, order_no="MO-ORD-%s" % code,
                production_dept_code=code, plan_quantity=Decimal("10"),
            )
            WorkOrderCostResult.objects.create(
                period=self.period, work_order=order, product_code="P-ORD",
                product_name="产品", plan_quantity=Decimal("10"),
                finished_quantity=Decimal("10"), wip_quantity=Decimal("0"),
                direct_material=Decimal("10"), finished_material=Decimal("10"),
                finished_cost=Decimal("10"), wip_cost=Decimal("0"),
                total_cost=Decimal("10"),
            )
        # 浆粕车间没有工单 → 用车间费用池让它出现（材料/完工按上传流水统计）
        WorkshopCostPool.objects.create(
            period=self.period, batch=batch, workshop="浆粕车间",
            water_power=Decimal("0"),
        )
        resp = self.client.get(self.url)
        self.assertEqual(
            [row["workshop"] for row in resp.context["rows"]],
            ["RanZheng", "tefang", "BaoFu", "ZhenZhi", "JiZhi", "浆粕车间"],
        )
        # 页面上第一行就是染色车间
        seg = resp.content.decode("utf-8").split("车间生产成本汇总", 1)[1]
        body = seg.split("<tbody>", 1)[1]
        self.assertLess(body.index("染色车间"), body.index("纺纱车间"))
        self.assertLess(body.index("纺纱车间"), body.index("包覆车间"))
        self.assertLess(body.index("包覆车间"), body.index("针织车间"))
        self.assertLess(body.index("针织车间"), body.index("机织车间"))
        self.assertLess(body.index("机织车间"), body.index("浆粕车间"))

    def test_export_and_print_sit_in_the_filter_rows(self):
        """导出Excel 在车间汇总筛选行内；三张表各有「打印本表」，都与筛选/清空同排。"""
        html = self.client.get(self.url).content.decode("utf-8")
        page_head = html.split("</section>", 1)[0]
        self.assertNotIn("导出Excel", page_head)      # 页头不再有导出
        forms = html.split('<form class="pc-filter-row"')[1:]
        self.assertEqual(len(forms), 3)
        summary_form = forms[0].split("</form>", 1)[0]
        for token in (">筛选<", ">清空<", "导出Excel", 'class="print-here"'):
            self.assertIn(token, summary_form, token)
        self.assertLess(summary_form.index(">清空<"), summary_form.index("导出Excel"))
        self.assertLess(
            summary_form.index("导出Excel"), summary_form.index('class="print-here"')
        )
        matrix_form = forms[1].split("</form>", 1)[0]
        for token in (">筛选<", ">清空<", 'class="print-here"'):
            self.assertIn(token, matrix_form, token)
        self.assertLess(
            matrix_form.index(">清空<"), matrix_form.index('class="print-here"')
        )
        # 速查表（用户 2026-09-24 新增）：同样有筛选/清空/打印本表
        issue_form = forms[2].split("</form>", 1)[0]
        for token in (">筛选<", ">清空<", 'class="print-here"'):
            self.assertIn(token, issue_form, token)
        self.assertLess(
            issue_form.index(">清空<"), issue_form.index('class="print-here"')
        )
        # 样式：这四个按钮与筛选项同高、底对齐
        style = html.split("<style>", 1)[1].split("</style>", 1)[0]
        self.assertIn(".pc-filter-row .filter-actions .button", style)
        self.assertIn("min-height:34px", style)
        self.assertIn("align-self:flex-end", style)

    def test_pulp_workshop_uses_uploaded_flow(self):
        """浆粕车间：材料与完工成本按上传的浆粕流水统计（不再只挂车间费用）。"""
        pulp_batch = self.create_batch(
            self.period, ImportBatch.Kind.PULP_FLOW, "pc-pulp"
        )
        product = Material.objects.create(
            code="21.03", name="芳纶浆粕DFP-103", uom="kg"
        )
        txn = self.create_txn(
            self.period, pulp_batch, txn_date=date(2026, 7, 20),
            document_no="PULP-OUT-1", business_type="生产入库",
            direction=InventoryTransaction.Direction.IN, material=product,
            quantity=Decimal("10"), order_no="PULP202607",
        )
        txn.amount = Decimal("0")
        txn.save(update_fields=["amount"])
        PulpMaterial.objects.create(
            period=self.period, batch=pulp_batch, txn_date=date(2026, 7, 5),
            document_no="PULP-IN-1", material_code="R52.101",
            material_name="对位芳纶短纤维", quantity=Decimal("8"),
            unit_price=Decimal("37.5"), amount=Decimal("300"),
            business_type="领料出库", uom="kg",
        )
        WorkshopCostPool.objects.create(
            period=self.period, batch=pulp_batch, workshop="浆粕车间",
            water_power=Decimal("100"), wages=Decimal("100"),
        )
        resp = self.client.get(self.url)
        rows = {row["workshop"]: row for row in resp.context["rows"]}
        self.assertIn("浆粕车间", rows)
        pulp = rows["浆粕车间"]
        # 材料＝原料领料金额；费用＝车间费用；完工成本＝材料＋费用（浆粕全完工、无在产品）
        self.assertEqual(pulp["current_material"], Decimal("300.00"))
        self.assertEqual(pulp["current_expenses"], Decimal("200.00"))
        self.assertEqual(pulp["current_input"], Decimal("500.00"))
        self.assertEqual(pulp["finished_cost"], Decimal("500.00"))
        self.assertEqual(pulp["ending_wip"], Decimal("0.00"))
        # 差额不再挂车间费用（以前是 200.00）
        self.assertEqual(pulp["difference"], Decimal("0.00"))
        self.assertTrue(pulp["is_calculated"])
        # 页面上也看得到 500.00（不再显示「待结转」）
        html = resp.content.decode("utf-8")
        seg = html.split("车间生产成本汇总", 1)[1].split("</table>", 1)[0]
        self.assertIn("500.00", seg)

    def test_old_total_cost_url_redirects(self):
        resp = self.client.get(reverse("total_cost", args=[self.period.id]))
        self.assertEqual(resp.status_code, 302)
        self.assertIn(reverse("production_cost_summary", args=[self.period.id]),
                      resp["Location"])
        resp2 = self.client.get(
            reverse("total_cost", args=[self.period.id]), {"workshop": "tefang"}
        )
        self.assertIn("workshop=tefang", resp2["Location"])

    def test_nav_has_single_entry(self):
        base = (Path(settings.BASE_DIR) / "templates" / "base.html").read_text(
            encoding="utf-8"
        )
        self.assertIn("{% url 'production_cost_summary' period.id %}", base)
        self.assertNotIn("{% url 'total_cost' period.id %}", base)
        self.assertFalse(
            (Path(settings.BASE_DIR) / "templates" / "settlement"
             / "total_cost.html").exists()
        )

    def test_export_has_two_sheets(self):
        resp = self.client.get(
            reverse("production_cost_export", args=[self.period.id])
        )
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(BytesIO(resp.content), read_only=True)
        try:
            self.assertEqual(wb.sheetnames, ["车间生产成本汇总", "成本构成表"])
            summary_sheet = wb["车间生产成本汇总"]
            heads = [c.value for c in next(summary_sheet.iter_rows(min_row=3, max_row=3))]
            self.assertEqual(
                heads,
                ["车间", "结转状态", "工单数", "期初在制", "本期材料", "本期费用",
                 "本期生产投入", "完工成本", "期末在制", "闭合差额"],
            )
            rows = list(summary_sheet.iter_rows(min_row=4, values_only=True))
            self.assertEqual(rows[0][0], "染色车间")
            self.assertEqual(rows[-1][0], "合计")
            matrix_sheet = wb["成本构成表"]
            first = next(matrix_sheet.iter_rows(min_row=1, max_row=1))[0].value
            self.assertEqual(first, "成本构成表")
        finally:
            wb.close()
        # 导出跟着筛选走（车间 / 期间）
        resp2 = self.client.get(
            reverse("production_cost_export", args=[self.period.id]),
            {"p_workshop": "tefang"},
        )
        self.assertEqual(resp2.status_code, 200)
        self.assertIn("tefang", resp2["Content-Disposition"])

    def test_old_export_url_redirects_to_new_one(self):
        resp = self.client.get(
            reverse("total_cost_export", args=[self.period.id])
        )
        self.assertEqual(resp.status_code, 302)
        self.assertIn(
            reverse("production_cost_export", args=[self.period.id]),
            resp["Location"],
        )


class EndingWipStatusAndProductTests(SettlementFixtureMixin, TestCase):
    """期末在产品页（用户 2026-09-21）：

    ①明细筛选行里的「在产品成本 xxx」金额徽标删除；
    ②加「工单状态」筛选（下拉，选完即时生效；导出同步）；
    ③产品编码 / 产品名称要显示出来：结算结果没带产品时，用 BOM 的「成品物料编号/名称」等兜底。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="ew-status", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.batch = self.create_batch(
            self.period, ImportBatch.Kind.WORK_ORDER, "ew-status"
        )
        # A：结算结果自带产品；B：结算结果不带产品，靠 BOM 兜底
        self.order_a = self._order("MO-EW-A", "RanZheng", "执行中")
        WorkOrderCostResult.objects.create(
            period=self.period, work_order=self.order_a, product_code="P-A",
            product_name="产品A", plan_quantity=Decimal("100"),
            finished_quantity=Decimal("40"), wip_quantity=Decimal("60"),
            wip_material=Decimal("600"), wip_wages=Decimal("100"),
            wip_overhead=Decimal("50"), wip_cost=Decimal("750"),
            finished_cost=Decimal("250"), total_cost=Decimal("1000"),
        )
        self.order_b = self._order("MO-EW-B", "tefang", "已完工")
        WorkOrderCostResult.objects.create(
            period=self.period, work_order=self.order_b, product_code="",
            product_name="", plan_quantity=Decimal("50"),
            finished_quantity=Decimal("10"), wip_quantity=Decimal("40"),
            wip_material=Decimal("400"), wip_wages=Decimal("0"),
            wip_overhead=Decimal("0"), wip_cost=Decimal("400"),
            finished_cost=Decimal("100"), total_cost=Decimal("500"),
        )
        BomLine.objects.create(
            period=self.period, batch=self.batch, work_order_no="MO-EW-B",
            material_code="P-B-BOM", material_name="BOM兜底产品",
            line_no="10", sub_item_code="R-1", sub_item_name="子项",
            qty_numerator=Decimal("1"), qty_denominator=Decimal("1"),
        )
        self.url = reverse("ending_wip", args=[self.period.id])

    def _order(self, order_no, workshop, status):
        return WorkOrder.objects.create(
            period=self.period, batch=self.batch, order_no=order_no,
            production_dept_code=workshop, status=status,
            plan_quantity=Decimal("100"),
        )

    def test_amount_pill_is_removed(self):
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertNotIn("detail-filter-pill", html)
        self.assertNotIn(">在产品成本 ", html)      # 筛选行里那枚金额徽标

    def test_status_filter_and_options(self):
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertIn('name="status"', html)
        # 下拉选项＝明细里出现过的状态
        seg = html.split('name="status"', 1)[1].split("</select>", 1)[0]
        self.assertIn("执行中", seg)
        self.assertIn("已完工", seg)
        resp = self.client.get(self.url, {"status": "已完工"})
        rows = resp.context["page"].object_list
        self.assertEqual([r["order_no"] for r in rows], ["MO-EW-B"])
        self.assertTrue(resp.context["wip_filter_active"])
        # 与车间组合
        resp2 = self.client.get(self.url, {"status": "已完工", "workshop": "RanZheng"})
        self.assertEqual(resp2.context["page"].object_list, [])
        # 导出跟着状态走
        exp = self.client.get(
            reverse("export_listing", args=[self.period.id, "ending_wip"]),
            {"status": "已完工"},
        )
        wb = load_workbook(BytesIO(exp.content), read_only=True)
        try:
            ws = wb[wb.sheetnames[0]]
            heads = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
            rows = list(ws.iter_rows(min_row=2, values_only=True))
        finally:
            wb.close()
        for column in ("工单状态", "产品编码", "产品名称"):
            self.assertIn(column, heads, column)
        # 导出末行是「合计」行，所以数据行 = 1 + 1
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0][heads.index("工单编号")], "MO-EW-B")
        self.assertEqual(rows[-1][0], "合计")

    def test_product_code_and_name_are_filled_from_bom(self):
        """结算结果没带产品时，用 BOM 的成品物料编号/名称兜底（原来显示 —）。"""
        resp = self.client.get(self.url)
        rows = {r["order_no"]: r for r in resp.context["page"].object_list}
        self.assertEqual(rows["MO-EW-A"]["product_code"], "P-A")
        self.assertEqual(rows["MO-EW-A"]["product_name"], "产品A")
        self.assertEqual(rows["MO-EW-B"]["product_code"], "P-B-BOM")
        self.assertEqual(rows["MO-EW-B"]["product_name"], "BOM兜底产品")
        # 页面上也真的显示出来（不再是 —）
        html = resp.content.decode("utf-8")
        body = html.split("<tbody>", 2)[-1]
        self.assertIn("P-B-BOM", body)
        self.assertIn("BOM兜底产品", body)
        # 产品编码筛选也能按兜底后的编码命中
        by_material = self.client.get(self.url, {"material": "P-B-BOM"})
        self.assertEqual(
            [r["order_no"] for r in by_material.context["page"].object_list],
            ["MO-EW-B"],
        )

class WorkshopFilterDropdownTests(SettlementFixtureMixin, TestCase):
    """车间下拉必须**始终列出全部车间**（用户 2026-09-21 报的 bug）。

    以前下拉是从「筛选后」的数据里取车间，所以选了纺纱之后下拉里就只剩纺纱，
    必须先切回「全部车间」才能换别的车间。现在下拉用完整的车间清单（与当前筛选无关）。
    覆盖：工单成本明细、生产成本的车间汇总 / 矩阵两张表。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="ws-opt", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        batch = self.create_batch(self.period, ImportBatch.Kind.WORK_ORDER, "ws-opt")
        for index, code in enumerate(
            ("RanZheng", "tefang", "BaoFu", "ZhenZhi", "JiZhi"), start=1
        ):
            order = WorkOrder.objects.create(
                period=self.period, batch=batch, order_no="MO-OPT-%d" % index,
                production_dept_code=code, plan_quantity=Decimal("10"),
            )
            WorkOrderCostResult.objects.create(
                period=self.period, work_order=order, product_code="P-OPT",
                product_name="产品", plan_quantity=Decimal("10"),
                finished_quantity=Decimal("10"), wip_quantity=Decimal("0"),
                direct_material=Decimal("10"), finished_material=Decimal("10"),
                finished_cost=Decimal("10"), wip_cost=Decimal("0"),
                total_cost=Decimal("10"),
            )
        WorkshopSettlement.objects.create(
            period=self.period, workshop="RanZheng", sequence=1,
            status=WorkshopSettlement.Status.SETTLED,
        )

    @staticmethod
    def _options(html, name):
        block = re.search(r'<select name="%s"[^>]*>(.*?)</select>' % name, html, re.S)
        return re.findall(r'<option value="([^"]*)"', block.group(1))

    def test_order_cost_dropdown_keeps_all_workshops(self):
        url = reverse("work_order_costs", args=[self.period.id])
        expected = ["", "RanZheng", "tefang", "BaoFu", "ZhenZhi", "JiZhi", "浆粕车间"]
        html = self.client.get(url).content.decode("utf-8")
        self.assertEqual(self._options(html, "w_workshop"), expected)
        # 选了纺纱之后，下拉仍然是完整的；并且能直接换成机织
        after = self.client.get(url, {"w_workshop": "tefang"})
        self.assertEqual(
            self._options(after.content.decode("utf-8"), "w_workshop"), expected
        )
        switched = self.client.get(url, {"w_workshop": "JiZhi"})
        self.assertEqual(switched.status_code, 200)
        self.assertEqual(
            {r.work_order.production_dept_code for r in switched.context["rows"]},
            {"JiZhi"},
        )

    def test_production_cost_dropdowns_keep_all_workshops(self):
        url = reverse("production_cost_summary", args=[self.period.id])
        expected = ["", "RanZheng", "tefang", "BaoFu", "ZhenZhi", "JiZhi", "浆粕车间"]
        html = self.client.get(url).content.decode("utf-8")
        self.assertEqual(self._options(html, "p_workshop"), expected)
        self.assertEqual(self._options(html, "m_workshop"), expected)
        # 车间汇总选了机织、矩阵选了纺纱：两张下拉都还是完整的
        resp = self.client.get(url, {"p_workshop": "JiZhi", "m_workshop": "tefang"})
        html2 = resp.content.decode("utf-8")
        self.assertEqual(self._options(html2, "p_workshop"), expected)
        self.assertEqual(self._options(html2, "m_workshop"), expected)
        # 两张表各自照样按自己的筛选收窄
        self.assertEqual(
            [row["workshop"] for row in resp.context["rows"]], ["JiZhi"]
        )
        self.assertEqual(
            [w["code"] for w in resp.context["total_cost"]["workshops"]], ["tefang"]
        )


class StandardCostPageTests(SettlementFixtureMixin, TestCase):
    """标准成本页（用户 2026-09-22 新增）：按车间录占比 + 列出各工单标准成本。

    口径：标准材料成本＝BOM「标准成品单价」×完工数量；标准人工/制造费用
    ＝标准材料成本×各自占比（基数＝材料成本）。
    用户 2026-09-22 二次确认：本页只出**标准侧**，实际成本以「工单成本」结果为准，
    单位标准成本放到「工单成本」页与实际单位成本对照。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="std-cost", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        batch = self.create_batch(self.period, ImportBatch.Kind.WORK_ORDER, "std-cost")
        # 染色：完工 10、标准成品单价 20（→标准材料 200）、实际材料 100/工资 20/费用 30
        self.order = WorkOrder.objects.create(
            period=self.period, batch=batch, order_no="MO-STD-1",
            production_dept_code="RanZheng", plan_quantity=Decimal("10"),
        )
        self.result = WorkOrderCostResult.objects.create(
            period=self.period, work_order=self.order, product_code="P-STD",
            product_name="标准成本测试产品", plan_quantity=Decimal("10"),
            finished_quantity=Decimal("10"), wip_quantity=Decimal("0"),
            direct_material=Decimal("100"), water_power=Decimal("10"),
            steam=Decimal("10"), depreciation=Decimal("10"),
            wages=Decimal("20"), total_cost=Decimal("150"),
            finished_material=Decimal("100"), finished_cost=Decimal("150"),
            wip_cost=Decimal("0"), finished_unit_cost=Decimal("15"),
        )
        BomLine.objects.create(
            period=self.period, batch=batch, work_order_no="MO-STD-1",
            material_code="P-STD", material_name="标准成本测试产品",
            sub_item_code="RM-STD", sub_item_name="原料", qty_numerator=Decimal("2"),
            qty_denominator=Decimal("1"), unit_price=Decimal("10"),
            standard_price=Decimal("20"),
        )

    def test_page_renders_rates_and_comparison(self):
        resp = self.client.get(reverse("standard_costs", args=[self.period.id]))
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        # 用户 2026-09-23：兜底参数整块已删除，六个车间不再有 labor_/overhead_ 输入框
        for code in ("RanZheng", "tefang", "BaoFu", "ZhenZhi", "JiZhi", "浆粕车间"):
            self.assertNotIn('name="labor_%s"' % code, html)
            self.assertNotIn('name="overhead_%s"' % code, html)
        self.assertNotIn("兜底参数", html)
        # 口径说明在表格之后（左下角）
        self.assertLess(html.index("</table>"), html.index('class="sc-basis"'))
        # 用户 2026-09-23 改名：页头＝标准成本；两张表＝标准成本构成表 / 工单标准成本计算表
        self.assertIn("<h1>标准成本</h1>", html)
        self.assertIn("<h2>标准成本构成表</h2>", html)
        self.assertIn("<h2>工单标准成本计算表</h2>", html)
        # 用户 2026-09-23 二次要求：构成表标题不带年份；期间下拉用「全部」；
        # 删掉表内分节行与「本月实际（参考）」列
        self.assertNotIn("标准成本构成表 · ", html)
        self.assertNotIn("全年（1-12 月）", html)
        self.assertIn('<option value="">全部</option>', html)
        self.assertNotIn("无需手工录", html)
        self.assertNotIn("本月实际（参考）", html)
        # 用户 2026-09-23 三次要求：删掉两句说明、删占比表的期间筛选、删加成系数解释句
        self.assertNotIn("会按半年滚动平均自动更新", html)
        self.assertNotIn("黄色底＝可录入", html)
        self.assertNotIn('name="r_month"', html)
        self.assertNotIn("＝ 1.7454 倍", html)
        # 占比表标题简短；「人工占比和费用占比的和」列在
        self.assertIn("适用的标准占比</h2>", html)
        self.assertIn("人工占比和费用占比的和", html)

    def test_composition_section_has_export_print_and_filter_buttons(self):
        """用户 2026-09-23：构成表筛选行里「筛选 / 清空筛选 / 导出Excel / 打印本表」对齐。"""
        html = self.client.get(
            reverse("standard_costs", args=[self.period.id])
        ).content.decode("utf-8")
        # 取构成表那一节里的筛选行（从 section 的 form 开始，别拿到 <style> 里的 CSS）
        start = html.index('<form class="sc-filter-row" method="get">')
        actions = html[html.index('<div class="filter-actions">', start):]
        actions = actions[:actions.index("</form>")]
        for label in ("筛选", "清空筛选", "导出Excel"):
            self.assertIn(label, actions)
        # 打印本表：筛选行里留了 print-here 容器（由 base.html 注入按钮）
        self.assertIn('class="print-here"', actions)
        # 四个动作在同一个按钮行里（对齐）
        self.assertLess(actions.index("筛选"), actions.index("清空筛选"))
        self.assertLess(actions.index("清空筛选"), actions.index("导出Excel"))
        self.assertLess(actions.index("导出Excel"), actions.index("print-here"))
        # 导出链接：带上构成表的期间/车间筛选参数
        url = reverse("standard_costs", args=[self.period.id])
        html = self.client.get(
            url, {"c_month": "2026-03", "c_workshop": "RanZheng"}
        ).content.decode("utf-8")
        start = html.index('<form class="sc-filter-row" method="get">')
        actions = html[html.index('<div class="filter-actions">', start):]
        actions = actions[:actions.index("</form>")]
        self.assertIn("c_month=2026-03", actions)
        self.assertIn("c_workshop=RanZheng", actions)

    def test_amount_and_pct_formatting(self):
        """用户 2026-09-23：金额加千分位、占比带 %。"""
        self.client.post(
            reverse("save_standard_cost_rates", args=[self.period.id]),
            {
                "qty_2026-01_RanZheng": "12345.6",
                "material_2026-01_RanZheng": "1000000",
                "wages_2026-01_RanZheng": "0",
                "overhead_2026-01_RanZheng": "0",
            },
        )
        url = reverse("standard_costs", args=[self.period.id])
        # 用户 2026-09-24：构成表默认只出当月，所以看 1 月的录入格要显式筛到 1 月
        html = self.client.get(url, {"c_month": "2026-01"}).content.decode("utf-8")
        # 录入格里金额带千分位（文本输入框，提交时后端会去掉逗号）
        self.assertIn('value="12,345.60"', html)
        self.assertIn('value="1,000,000.00"', html)
        # 占比格子是自动算出来的文本 + 百分号（1 月：材料 100%）
        self.assertIn('data-pct="material">100.00%</td>', html)
        # 下半年（有实际数）的占比带百分号：7 月染色 材料占比 = 100÷150 = 66.67%
        html = self.client.get(url).content.decode("utf-8")
        july_block = html[html.index("2026-07"):html.index("适用的标准占比")]
        self.assertIn("66.67%", july_block)
        self.assertIn("13.33%", july_block)   # 人工 = 20÷150

    def test_composition_amounts_are_the_only_inputs(self):
        """用户 2026-09-23：上半年只填金额，占比不再手填（没有占比输入框）。

        用户 2026-09-24：构成表默认只显示当月，所以这里显式筛到 1 月看黄底录入格。
        """
        html = self.client.get(
            reverse("standard_costs", args=[self.period.id]), {"c_month": "2026-01"}
        ).content.decode("utf-8")
        self.assertNotIn('name="material_pct_', html)
        self.assertNotIn('name="labor_pct_', html)
        self.assertNotIn('name="overhead_pct_', html)
        # 四个金额字段都在（1 月染色为例）
        for field in ("qty", "material", "wages", "overhead"):
            self.assertIn('name="%s_2026-01_RanZheng"' % field, html)

    def test_composition_defaults_to_current_month(self):
        """用户 2026-09-24：标准成本构成表默认显示当月份的数据。

        - 不带 c_month 参数 → 只出当月的行（不再是 1-12 月 72 行）；
        - c_month 显式传空（下拉选「全部」）→ 仍然一年 12 个月。
        """
        url = reverse("standard_costs", args=[self.period.id])
        resp = self.client.get(url)
        rows = resp.context["composition_rows"]
        self.assertEqual(sorted({r["period_key"] for r in rows}), ["2026-07"])
        html = resp.content.decode("utf-8")
        # 下拉里当月的选项被选中（默认值＝当月）
        self.assertIn('<option value="2026-07" selected>2026-07</option>', html)
        self.assertEqual(resp.context["c_filter"]["month"], "2026-07")
        # 下拉选「全部」→ 一年 12 个月 × 6 车间
        resp = self.client.get(url, {"c_month": ""})
        self.assertEqual(
            sorted({r["period_key"] for r in resp.context["composition_rows"]}),
            ["2026-%02d" % m for m in range(1, 13)],
        )
        self.assertEqual(resp.context["c_filter"]["month"], "")

    def test_rate_uses_average_of_monthly_ratios(self):
        """用户 2026-09-23：占比＝基期各月占比的平均数（不是半年合计数）。"""
        from settlement.models import StandardCostCompositionEntry

        # 1 月材料小、费用占比高；2 月材料大、费用占比低 —— 两种算法结果不同
        for month, (m, l, o) in (
            (date(2026, 1, 1), (100, 50, 50)),
            (date(2026, 2, 1), (900, 90, 90)),
        ):
            StandardCostCompositionEntry.objects.create(
                period_month=month, workshop="RanZheng", output_qty=Decimal("1"),
                direct_material=Decimal(m), wages=Decimal(l),
                manufacturing_overhead=Decimal(o),
            )
        _rows, meta = views.build_standard_cost_rows(self.period)
        rate = meta["rates"]["RanZheng"]

        def q(value):
            return Decimal(value).quantize(Decimal("0.000001"))

        # 各月占比：1 月 50/25/25%；2 月 83.3333/8.3333/8.3333% → 取平均
        # （若按半年合计数会得到 材料 78.125%、人工 10.9375%，能区分开）
        self.assertEqual(rate["source"], "rolling")
        self.assertEqual(rate["months"], 2)
        self.assertEqual(q(rate["material_pct"]), Decimal("66.666667"))
        self.assertEqual(q(rate["labor_pct"]), Decimal("16.666667"))
        self.assertEqual(q(rate["overhead_pct"]), Decimal("16.666667"))
        # 「人工占比和费用占比的和」＝人工＋制造费用
        self.assertEqual(q(rate["labor_plus_overhead_pct"]), Decimal("33.333333"))
        # 加成系数 = 1 + 16.6667% + 16.6667% ≈ 1.3333
        self.assertEqual(
            q(views.standard_cost_factor(meta["rates"], "RanZheng")),
            Decimal("1.333333"),
        )

    def test_rate_table_shows_labor_plus_overhead_column(self):
        """用户 2026-09-23：「适用的标准占比」表删掉三列，加「人工占比和费用占比的和」。"""
        html = self.client.get(
            reverse("standard_costs", args=[self.period.id])
        ).content.decode("utf-8")
        block = html[html.index("适用的标准占比"):html.index("工单标准成本计算表")]
        self.assertIn("人工占比和费用占比的和", block)
        # 删掉的三列不再出现
        self.assertNotIn("人工(占材料)%", block)
        self.assertNotIn("制造费用(占材料)%", block)
        self.assertNotIn("标准加成系数", block)
        # 表下不再有「怎么算」示例句（用户 2026-09-23 要求删除）
        self.assertNotIn("怎么算：", block)

    def test_bom_qty_over_100_is_scaled_to_100(self):
        """用户 2026-09-23：BOM 子项用量合计 > 100% 时，按用量比例缩放到 100% 再算标准成品单价。

        规则：Σ用量 > 1 → 标准成品单价 ＝ Σ(用量×单价) ÷ Σ用量；Σ用量 ≤ 1 的不动。
        """
        from settlement.models import InventoryValuation
        from settlement.services.costing import _fill_bom_prices

        batch = ImportBatch.objects.create(
            period=self.period, kind=ImportBatch.Kind.BOM, original_name="bom-norm"
        )
        for code, price in (("RM-A", Decimal("10")), ("RM-B", Decimal("40"))):
            InventoryValuation.objects.create(
                period=self.period, material_code=code,
                weighted_average_price=price,
            )
        # 单 1：Σ用量 = 3（1 + 2）→ 应缩放：Σ(用量×单价)=1×10+2×40=90 → 90÷3=30
        for code, num in (("RM-A", Decimal("1")), ("RM-B", Decimal("2"))):
            BomLine.objects.create(
                period=self.period, batch=batch, work_order_no="MO-OVER",
                material_code="P-OVER", material_name="超100%产品",
                sub_item_code=code, qty_numerator=num, qty_denominator=Decimal("1"),
            )
        # 单 2：Σ用量 = 1（0.5 + 0.5）→ 不动：Σ(用量×单价)=0.5×10+0.5×40=25
        for code, num in (("RM-A", Decimal("50")), ("RM-B", Decimal("50"))):
            BomLine.objects.create(
                period=self.period, batch=batch, work_order_no="MO-EXACT",
                material_code="P-EXACT", material_name="正好100%产品",
                sub_item_code=code, qty_numerator=num, qty_denominator=Decimal("100"),
            )
        # 单 3：Σ用量 = 0.6（0.2 + 0.4）→ ≤1 不动：0.2×10+0.4×40=18
        for code, num in (("RM-A", Decimal("20")), ("RM-B", Decimal("40"))):
            BomLine.objects.create(
                period=self.period, batch=batch, work_order_no="MO-UNDER",
                material_code="P-UNDER", material_name="不足100%产品",
                sub_item_code=code, qty_numerator=num, qty_denominator=Decimal("100"),
            )

        _fill_bom_prices(self.period)

        def std(order):
            return BomLine.objects.filter(work_order_no=order).first().standard_price

        self.assertEqual(std("MO-OVER"), Decimal("30.00"))    # 90 ÷ 3
        self.assertEqual(std("MO-EXACT"), Decimal("25.00"))   # 不缩放
        self.assertEqual(std("MO-UNDER"), Decimal("18.00"))   # ≤100% 不缩放
        # 子项单价仍是月末加权价（不参与缩放）
        self.assertEqual(
            BomLine.objects.get(work_order_no="MO-OVER", sub_item_code="RM-B").unit_price,
            Decimal("40.00000000"),
        )

    def test_bom_unit_price_falls_back_to_historical_price(self):
        """用户 2026-09-30：BOM 子项单价 = 月末加权价 → **历史单价（出库价优先、入库价兜底）**。

        以前只查月末加权价，没结转到的物料一律记 0（BOM 表一大堆没单价）；
        现在与投料单价/BOM 反推同一口径。
        """
        from settlement.models import HistoricalPrice, InventoryValuation
        from settlement.services.costing import _fill_bom_prices

        batch = ImportBatch.objects.create(
            period=self.period, kind=ImportBatch.Kind.BOM, original_name="bom-price"
        )
        # RM-VAL：有月末加权价（优先）
        InventoryValuation.objects.create(
            period=self.period, material_code="RM-VAL",
            weighted_average_price=Decimal("10"),
        )
        # RM-HIS：没有月末价，只有历史单价（出库价优先于入库价）
        HistoricalPrice.objects.create(
            period=self.period, material_code="RM-HIS",
            issue_price=Decimal("61.95"), receipt_price=Decimal("50"),
        )
        # RM-RCP：只有入库价
        HistoricalPrice.objects.create(
            period=self.period, material_code="RM-RCP", receipt_price=Decimal("7.5"),
        )
        # RM-NONE：哪儿都没有价
        for code in ("RM-VAL", "RM-HIS", "RM-RCP", "RM-NONE"):
            BomLine.objects.create(
                period=self.period, batch=batch, work_order_no="MO-PRICE",
                material_code="P-PRICE", sub_item_code=code,
                qty_numerator=Decimal("1"), qty_denominator=Decimal("1"),
            )

        _fill_bom_prices(self.period)

        def unit(code):
            return BomLine.objects.get(sub_item_code=code).unit_price

        self.assertEqual(unit("RM-VAL"), Decimal("10.00000000"))     # 月末加权价优先
        self.assertEqual(unit("RM-HIS"), Decimal("61.95000000"))     # 历史出库价
        self.assertEqual(unit("RM-RCP"), Decimal("7.50000000"))      # 历史入库价兜底
        self.assertEqual(unit("RM-NONE"), Decimal("0.00000000"))     # 两处都没有 → 0
        # 标准成品单价 = Σ(用量×单价) ＝ 10 + 61.95 + 7.5 + 0（Σ用量 = 4 > 1 → 折算 ÷4）
        self.assertEqual(
            BomLine.objects.filter(work_order_no="MO-PRICE").first().standard_price,
            Decimal("19.86"),
        )

    def test_calc_table_has_period_filter(self):
        """用户 2026-09-23：工单标准成本计算表加「期间」筛选字段。"""
        june = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        url = reverse("standard_costs", args=[self.period.id])
        resp = self.client.get(url)
        html = resp.content.decode("utf-8")
        block = html[html.index("<h2>工单标准成本计算表</h2>"):]
        self.assertIn('name="w_period"', block[:1500])
        # 默认本期（7 月）
        self.assertIn("本期（2026-07）", block[:1500])
        # 选到 6 月 → 页面按 6 月出数（基期变成 2025 下半年）
        resp = self.client.get(url, {"w_period": "2026-06"})
        self.assertEqual(resp.context["period"].id, june.id)
        self.assertEqual(resp.context["period_label"], "2026-06")
        self.assertEqual(resp.context["view_basis_label"], "2025年下半年")
        # 构成表也跟着走那一期（用户 2026-09-24：默认只出当期那个月）
        self.assertEqual(
            sorted({r["period_key"] for r in resp.context["composition_rows"]}),
            ["2026-06"],
        )

    def test_period_filter_for_rolling_rates(self):
        """用户 2026-09-23：占比表不再单独筛选期（跟本页期间走）。"""
        from settlement.models import StandardCostCompositionEntry

        StandardCostCompositionEntry.objects.create(
            period_month=date(2026, 3, 1), workshop="RanZheng",
            direct_material=Decimal("70"), wages=Decimal("20"),
            manufacturing_overhead=Decimal("10"),
        )
        url = reverse("standard_costs", args=[self.period.id])
        resp = self.client.get(url)
        self.assertEqual(resp.context["rate_month_label"], "2026-07")
        self.assertEqual(resp.context["view_basis_label"], "2026年上半年")
        self.assertNotIn('name="r_month"', resp.content.decode("utf-8"))
        # 切到别的期间（下一页面的期间筛选）→ 基期跟着变
        AccountingPeriod.objects.create(month=date(2027, 3, 1))
        resp = self.client.get(url, {"w_period": "2027-03"})
        self.assertEqual(resp.context["view_basis_label"], "2026年下半年")
        row = next(
            r for r in resp.context["standard_rate_rows"] if r["workshop"] == "RanZheng"
        )
        self.assertEqual(row["months"], 1)          # 2026-07 只有一个月有实际数
        self.assertEqual(row["source"], "rolling")

    def test_ajax_save_returns_json_and_new_rates(self):
        """用户 2026-09-23：点「保存构成表」不跳页（返回 JSON + 新的半年平均占比）。"""
        resp = self.client.post(
            reverse("save_standard_cost_rates", args=[self.period.id]),
            {
                "qty_2026-01_RanZheng": "1",
                "material_2026-01_RanZheng": "60",
                "wages_2026-01_RanZheng": "30",
                "overhead_2026-01_RanZheng": "10",
            },
        )
        # 不是重定向，而是 JSON
        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp["Content-Type"], "application/json")
        payload = resp.json()
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["basis_label"], "2026年上半年")
        rate = next(r for r in payload["rates"] if r["workshop"] == "RanZheng")
        self.assertEqual(Decimal(rate["material_pct"]).quantize(Decimal("0.01")),
                         Decimal("60.00"))
        self.assertEqual(Decimal(rate["labor_pct"]).quantize(Decimal("0.01")),
                         Decimal("30.00"))
        self.assertEqual(rate["months"], 1)
        self.assertEqual(rate["source"], "rolling")

    def test_save_composition_clears_emptied_rows(self):
        """四个金额都清空 = 删掉这条录入（避免旧数据留着骗人）。"""
        from settlement.models import StandardCostCompositionEntry

        url = reverse("save_standard_cost_rates", args=[self.period.id])
        self.client.post(url, {
            "qty_2026-01_RanZheng": "1", "material_2026-01_RanZheng": "60",
            "wages_2026-01_RanZheng": "30", "overhead_2026-01_RanZheng": "10",
        })
        self.assertTrue(StandardCostCompositionEntry.objects.exists())
        self.client.post(url, {
            "qty_2026-01_RanZheng": "", "material_2026-01_RanZheng": "",
            "wages_2026-01_RanZheng": "", "overhead_2026-01_RanZheng": "",
        })
        self.assertFalse(StandardCostCompositionEntry.objects.exists())

    def test_redirect_entry_still_works(self):
        """兜底参数表单走跳转入口：保存后回到页面并给 messages 提示。"""
        resp = self.client.post(
            reverse("save_standard_cost_rates_redirect", args=[self.period.id]),
            {"labor_RanZheng": "20", "overhead_RanZheng": "30"},
            follow=True,
        )
        self.assertEqual(resp.status_code, 200)
        rate = StandardCostRate.objects.get(period=self.period, workshop="RanZheng")
        self.assertEqual(rate.labor_rate_pct, Decimal("20.0000"))
        self.assertEqual(rate.overhead_rate_pct, Decimal("30.0000"))

    def test_composition_table_columns(self):
        """用户 2026-09-23：构成表 9 列（含人工；期间×车间逐月列出）。"""
        html = self.client.get(
            reverse("standard_costs", args=[self.period.id])
        ).content.decode("utf-8")
        block = html[html.index("标准成本构成表"):html.index("工单标准成本计算表")]
        heads = [
            "期间", "车间", "产量", "直接材料", "人工", "制造费用",
            "材料成本占比%", "人工成本占比%", "制造费用占比%",
        ]
        last = -1
        for name in heads:
            idx = block.index(">%s<" % name, last + 1)
            self.assertGreater(idx, last, "「%s」列顺序不对" % name)
            last = idx
        # 期间/车间筛选字段
        self.assertIn('name="c_month"', block)
        self.assertIn('name="c_workshop"', block)

    def test_composition_rows_cover_the_year(self):
        """构成表＝参考年 1-12 月 × 车间；7 月起自动取实际发生数。

        用户 2026-09-24：默认只出当月，所以这里显式选「全部」（c_month 传空）看整年。
        """
        resp = self.client.get(
            reverse("standard_costs", args=[self.period.id]), {"c_month": ""}
        )
        rows = resp.context["composition_rows"]
        months = sorted({r["period_key"] for r in rows})
        self.assertEqual(months, ["2026-%02d" % m for m in range(1, 13)])
        # 上半年可录入、下半年系统自动
        self.assertEqual(
            sorted({r["period_key"] for r in rows if r["editable"]}),
            ["2026-%02d" % m for m in range(1, 7)],
        )
        self.assertFalse(
            any(r["editable"] for r in rows if r["period_key"] >= "2026-07")
        )
        july = next(
            r for r in rows if r["period_key"] == "2026-07" and r["workshop"] == "RanZheng"
        )
        # 7 月：完工 10、材料 100、人工 20、制造费用 10+10+10=30 → 合计 150
        self.assertEqual(july["output_qty"], Decimal("10"))
        self.assertEqual(july["direct_material"], Decimal("100"))
        self.assertEqual(july["wages"], Decimal("20"))
        self.assertEqual(july["manufacturing_overhead"], Decimal("30"))
        self.assertEqual(july["material_pct"], Decimal("100") / Decimal("150") * 100)
        self.assertEqual(july["labor_pct"], Decimal("20") / Decimal("150") * 100)
        self.assertEqual(july["overhead_pct"], Decimal("30") / Decimal("150") * 100)
        self.assertEqual(july["source"], "actual")
        # 8 月没有数据 → 空
        august = next(
            r for r in rows if r["period_key"] == "2026-08" and r["workshop"] == "RanZheng"
        )
        self.assertIsNone(august["output_qty"])

    def test_composition_filter_by_month_and_workshop(self):
        url = reverse("standard_costs", args=[self.period.id])
        resp = self.client.get(url, {"c_month": "2026-03", "c_workshop": "RanZheng"})
        rows = resp.context["composition_rows"]
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["period_key"], "2026-03")
        self.assertEqual(rows[0]["workshop"], "RanZheng")

    def test_save_manual_composition_from_post(self):
        """上半年手工录入金额 → 占比自动算（用户 2026-09-23）。"""
        from settlement.models import StandardCostCompositionEntry

        resp = self.client.post(
            reverse("save_standard_cost_rates", args=[self.period.id]),
            {
                "qty_2026-01_RanZheng": "12",
                "material_2026-01_RanZheng": "60",
                "wages_2026-01_RanZheng": "30",
                "overhead_2026-01_RanZheng": "10",
            },
        )
        self.assertEqual(resp.status_code, 200)
        self.assertTrue(resp.json()["ok"])
        entry = StandardCostCompositionEntry.objects.get(
            period_month=date(2026, 1, 1), workshop="RanZheng"
        )
        self.assertEqual(entry.output_qty, Decimal("12"))
        self.assertEqual(entry.direct_material, Decimal("60"))
        resp = self.client.get(
            reverse("standard_costs", args=[self.period.id]), {"c_month": "2026-01"}
        )
        row = next(
            r for r in resp.context["composition_rows"]
            if r["period_key"] == "2026-01" and r["workshop"] == "RanZheng"
        )
        self.assertEqual(row["material_pct"], Decimal("60"))
        self.assertEqual(row["labor_pct"], Decimal("30"))
        self.assertEqual(row["overhead_pct"], Decimal("10"))

    def test_half_year_rolling_rate_drives_standard_cost(self):
        """用户 2026-09-23：标准占比＝上一个完整半年的实际占比平均（每半年滚动）。"""
        from settlement.models import StandardCostCompositionEntry

        # 2026 上半年两个月：平均后 材料 50%、人工 20%、制造费用 30%
        for month, values in (
            (date(2026, 1, 1), (60, 30, 10)),
            (date(2026, 2, 1), (40, 10, 50)),
        ):
            StandardCostCompositionEntry.objects.create(
                period_month=month, workshop="RanZheng", output_qty=Decimal("1"),
                direct_material=Decimal(values[0]), wages=Decimal(values[1]),
                manufacturing_overhead=Decimal(values[2]),
            )
        rows, meta = views.build_standard_cost_rows(self.period)
        self.assertEqual(meta["basis_label"], "2026年上半年")
        rate = meta["rates"]["RanZheng"]
        self.assertEqual(rate["source"], "rolling")
        self.assertEqual(rate["months"], 2)
        self.assertEqual(rate["material_pct"], Decimal("50"))
        self.assertEqual(rate["labor_pct"], Decimal("20"))
        self.assertEqual(rate["overhead_pct"], Decimal("30"))
        row = next(r for r in rows if r["order_no"] == "MO-STD-1")
        # 标准材料 20×10=200；人工 200×20%=40；制造费用 200×30%=60；合计 300
        self.assertEqual(row["standard_material"], Decimal("200"))
        self.assertEqual(row["standard_labor"], Decimal("40"))
        self.assertEqual(row["standard_overhead"], Decimal("60"))
        self.assertEqual(row["standard_total"], Decimal("300"))
        self.assertEqual(row["standard_unit_cost"], Decimal("30"))

    def test_rolling_rate_skips_months_without_data(self):
        """基期里只有一个月有数时，平均就用那一个月；没数据的月不拉低平均。"""
        from settlement.models import StandardCostCompositionEntry

        StandardCostCompositionEntry.objects.create(
            period_month=date(2026, 4, 1), workshop="RanZheng",
            direct_material=Decimal("25"), wages=Decimal("50"),
            manufacturing_overhead=Decimal("25"),
        )
        _rows, meta = views.build_standard_cost_rows(self.period)
        rate = meta["rates"]["RanZheng"]
        self.assertEqual(rate["months"], 1)
        self.assertEqual(rate["material_pct"], Decimal("25"))
        self.assertEqual(rate["labor_pct"], Decimal("50"))

    def test_calculation_table_has_period_column(self):
        """用户 2026-09-23：工单标准成本计算表加「期间」列（第一列）。

        用户 2026-09-24：删掉「完工数量 / 材料单位成本 / 总成本」三列，
        「材料/人工/制造费用」改名为「标准材料成本/标准人工成本/标准制造费用」，
        并且「标准成本」这个分组表头居中。
        """
        html = self.client.get(
            reverse("standard_costs", args=[self.period.id])
        ).content.decode("utf-8")
        block = html[html.index("工单标准成本计算表"):html.index("<h2>BOM 明细表</h2>")]
        self.assertLess(block.index(">期间<"), block.index(">车间<"))
        self.assertIn(">2026-07<", block)
        # 新列名
        for name in ("标准材料成本", "标准人工成本", "标准制造费用", "单位成本"):
            self.assertIn(">%s<" % name, block)
        # 删掉的三列不再出现
        self.assertNotIn(">完工数量<", block)
        self.assertNotIn(">材料单位成本<", block)
        self.assertNotIn(">总成本<", block)
        # 「标准成本」分组表头居中（跨 4 列）
        self.assertIn('<th colspan="4" class="sc-group-std sc-group-title">标准成本</th>', block)
        self.assertIn(".sc-group-title{text-align:center;}", html)

    def test_rates_save_and_feed_the_numbers(self):
        resp = self.client.post(
            reverse("save_standard_cost_rates", args=[self.period.id]),
            {"labor_RanZheng": "20", "overhead_RanZheng": "30"},
            follow=True,
        )
        self.assertEqual(resp.status_code, 200)
        rate = StandardCostRate.objects.get(period=self.period, workshop="RanZheng")
        self.assertEqual(rate.labor_rate_pct, Decimal("20.0000"))
        self.assertEqual(rate.overhead_rate_pct, Decimal("30.0000"))

        rows, meta = views.build_standard_cost_rows(self.period)
        row = next(r for r in rows if r["order_no"] == "MO-STD-1")
        # 标准材料 = 20 × 10 = 200；人工 = 200×20% = 40；制造费用 = 200×30% = 60
        self.assertEqual(row["standard_material"], Decimal("200"))
        self.assertEqual(row["standard_labor"], Decimal("40"))
        self.assertEqual(row["standard_overhead"], Decimal("60"))
        self.assertEqual(row["standard_total"], Decimal("300"))
        self.assertEqual(row["standard_unit_cost"], Decimal("30"))
        # 只算标准侧：实际成本与差异在「工单成本」页对照，这里不再重复算
        self.assertNotIn("actual_total", row)
        self.assertNotIn("diff_amount", row)

    def test_invalid_rate_is_rejected(self):
        url = reverse("save_standard_cost_rates_redirect", args=[self.period.id])
        for bad in ("abc", "-5", "99999"):
            resp = self.client.post(url, {"labor_RanZheng": bad}, follow=True)
            self.assertEqual(resp.status_code, 200)
            self.assertFalse(
                StandardCostRate.objects.filter(
                    period=self.period, workshop="RanZheng"
                ).exists(),
                "非法输入 %s 被写进库了" % bad,
            )

    def test_invalid_amount_is_reported_without_crash(self):
        """非法输入走 AJAX：返回 ok=False + 中文提示（不是 500）。"""
        resp = self.client.post(
            reverse("save_standard_cost_rates", args=[self.period.id]),
            {"qty_2026-01_RanZheng": "1", "material_2026-01_RanZheng": "abc"},
        )
        self.assertEqual(resp.status_code, 200)
        payload = resp.json()
        self.assertFalse(payload["ok"])
        self.assertIn("不是数字", payload["message"])

    def test_rates_carry_over_when_current_period_is_empty(self):
        # 6 月录过、7 月没录 → 7 月沿用 6 月的占比
        june = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        StandardCostRate.objects.create(
            period=june, workshop="RanZheng",
            labor_rate_pct=Decimal("10"), overhead_rate_pct=Decimal("20"),
        )
        rows, meta = views.build_standard_cost_rows(self.period)
        row = next(r for r in rows if r["order_no"] == "MO-STD-1")
        self.assertEqual(meta["carried_from"], june)
        self.assertEqual(row["standard_labor"], Decimal("20"))     # 200 × 10%
        self.assertEqual(row["standard_overhead"], Decimal("40"))  # 200 × 20%

    def test_order_without_bom_has_dash(self):
        # 标准价算不出的行：标准侧与差异都是空（None）
        batch = self.create_batch(self.period, ImportBatch.Kind.WORK_ORDER, "std-nobom")
        order = WorkOrder.objects.create(
            period=self.period, batch=batch, order_no="MO-STD-NOBOM",
            production_dept_code="JiZhi", plan_quantity=Decimal("5"),
        )
        WorkOrderCostResult.objects.create(
            period=self.period, work_order=order, product_code="P-NOBOM",
            product_name="无BOM产品", plan_quantity=Decimal("5"),
            finished_quantity=Decimal("5"), wip_quantity=Decimal("0"),
            direct_material=Decimal("50"), total_cost=Decimal("50"),
            finished_material=Decimal("50"), finished_cost=Decimal("50"),
            wip_cost=Decimal("0"), finished_unit_cost=Decimal("10"),
        )
        rows, meta = views.build_standard_cost_rows(self.period)
        row = next(r for r in rows if r["order_no"] == "MO-STD-NOBOM")
        self.assertFalse(row["has_standard_price"])
        self.assertIsNone(row["standard_material"])
        self.assertIsNone(row["standard_unit_cost"])
        self.assertEqual(meta["no_standard_price"], 1)

    def test_bom_with_zero_standard_price_counts_as_unpriced(self):
        """BOM 行存在、但标准价算成 0（子项原料当月没价）时，按「算不出」处理。

        用户 2026-09-22 核实：页面上原先写「无 BOM」是错的，实际是 BOM 有、价为 0。
        """
        BomLine.objects.create(
            period=self.period, batch=self.period.import_batches.first(),
            work_order_no="MO-STD-1", material_code="P-STD",
            sub_item_code="RM-ZERO", qty_numerator=Decimal("1"),
            qty_denominator=Decimal("1"), unit_price=Decimal("0"),
            standard_price=Decimal("0"),
        )
        # 把原来那条有价 BOM 改成 0 → 整单标准价 = 0
        BomLine.objects.filter(
            period=self.period, work_order_no="MO-STD-1", sub_item_code="RM-STD"
        ).update(standard_price=Decimal("0"))
        rows, meta = views.build_standard_cost_rows(self.period)
        row = next(r for r in rows if r["order_no"] == "MO-STD-1")
        self.assertFalse(row["has_standard_price"])
        self.assertIsNone(row["standard_material"])
        self.assertEqual(meta["no_standard_price"], 1)
        # 页面上措辞不能再说「无 BOM」；用户 2026-09-23 要求标题不再带张数说明，
        # 所以「标准成本算不出」这句只保留在合计行（用「标准成本为 —」表述）。
        html = self.client.get(
            reverse("standard_costs", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertNotIn("张无 BOM", html)
        self.assertIn("标准成本为 —", html)
        # 标题＝工单标准成本计算表（不带期间/张数后缀）
        self.assertIn("<h2>工单标准成本计算表</h2>", html)
        self.assertNotIn("工单标准成本计算表 · ", html)

    def test_export_has_two_sheets_and_totals_row(self):
        self.client.post(
            reverse("save_standard_cost_rates", args=[self.period.id]),
            {"labor_RanZheng": "20", "overhead_RanZheng": "30"},
        )
        resp = self.client.get(
            reverse("export_listing", args=[self.period.id, "standard_cost"])
        )
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(BytesIO(resp.content))
        # 用户 2026-09-23：加第三张「BOM 明细」（明细见 test_standard_cost_export_has_bom_detail_sheet）
        self.assertEqual(
            wb.sheetnames,
            ["工单标准成本计算表", "标准成本构成表", "BOM 明细"],
        )
        sheet = wb["工单标准成本计算表"]
        head = [cell.value for cell in sheet[1]]
        self.assertEqual(head[0], "期间")
        # 用户 2026-09-24：删掉完工数量/标准材料单位成本/标准总成本，三列改名
        self.assertEqual(head[5], "标准材料成本")
        self.assertEqual(head[6], "标准人工成本")
        self.assertEqual(head[7], "标准制造费用")
        self.assertEqual(head[8], "标准单位成本")
        self.assertNotIn("完工数量", head)
        self.assertNotIn("标准材料单位成本", head)
        self.assertNotIn("标准总成本", head)
        # 用户 2026-09-22：不重复实际成本与差异（那在「工单成本」里）
        self.assertNotIn("实际总成本", head)
        self.assertNotIn("差异额（实际−标准）", head)
        # 合计行与表头同列对齐：第 6/7/8 列＝材料/人工/制造费用，第 9 列＝标准单位成本
        last = [cell.value for cell in sheet[sheet.max_row]]
        self.assertEqual(Decimal(str(last[5])), Decimal("200"))
        self.assertEqual(Decimal(str(last[6])), Decimal("40"))
        self.assertEqual(Decimal(str(last[7])), Decimal("60"))
        self.assertEqual(Decimal(str(last[8])), Decimal("30"))
        # 第二张：标准成本构成表（用户 2026-09-23：期间×车间逐月 + 标准占比汇总）
        params = wb["标准成本构成表"]
        self.assertEqual(params["A1"].value, "期间")
        self.assertEqual(params["B1"].value, "车间")
        self.assertEqual(params["E1"].value, "人工")
        self.assertEqual(params["F1"].value, "制造费用")
        self.assertEqual(params["G1"].value, "材料成本占比%")
        # 第一行数据＝2026-01 染色车间（全年逐月，1 月在前）
        self.assertEqual(params["A2"].value, "2026-01")
        self.assertEqual(params["B2"].value, "染色车间")
        # 本年 1-12 月 × 6 个车间 + 空行 + 标准占比汇总
        self.assertEqual(params["A2"].value, "2026-01")
        values = [params.cell(row=i, column=1).value for i in range(1, params.max_row + 1)]
        self.assertIn("标准占比", values)
        # 兜底参数（labor_/overhead_）仍然是按百分数存
        self.client.post(
            reverse("save_standard_cost_rates", args=[self.period.id]),
            {"labor_RanZheng": "20", "overhead_RanZheng": "30"},
        )
        rate = StandardCostRate.objects.get(period=self.period, workshop="RanZheng")
        self.assertEqual(rate.labor_rate_pct, Decimal("20.0000"))
        self.assertEqual(rate.overhead_rate_pct, Decimal("30.0000"))

    def test_export_respects_composition_filters(self):
        """用户 2026-09-23：构成表的「期间/车间」筛选同样作用到导出。"""
        url = reverse("export_listing", args=[self.period.id, "standard_cost"])
        # 全年全车间：12 个月 × 6 个车间（+ 汇总段）
        wb = load_workbook(BytesIO(self.client.get(url).content))
        self.assertEqual(wb["标准成本构成表"]["A2"].value, "2026-01")
        # 只看 2026-07 + 染色车间
        wb = load_workbook(
            BytesIO(self.client.get(url, {"c_month": "2026-07", "c_workshop": "RanZheng"}).content)
        )
        sheet = wb["标准成本构成表"]
        self.assertEqual(sheet["A2"].value, "2026-07")
        self.assertEqual(sheet["B2"].value, "染色车间")
        self.assertEqual(sheet["C2"].value, 10)          # 产量
        self.assertAlmostEqual(float(sheet["G2"].value), 66.6666666, places=5)
        # 第二行就是汇总段（只剩一个车间一个月）
        self.assertEqual(sheet["A3"].value, None)

    def test_work_order_cost_page_shows_standard_unit_cost(self):
        """用户 2026-09-22：标准单位成本放到「工单成本」页，与实际单位成本对照。"""
        self.client.post(
            reverse("save_standard_cost_rates", args=[self.period.id]),
            {"labor_RanZheng": "20", "overhead_RanZheng": "30"},
        )
        url = reverse("work_order_costs", args=[self.period.id])
        resp = self.client.get(url)
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        # 两列加在「单位成本」之后
        self.assertIn('<th class="num">标准单位成本</th>', html)
        self.assertIn('<th class="num">单位成本差异</th>', html)
        self.assertLess(html.index(">单位成本<"), html.index(">标准单位成本<"))
        self.assertLess(html.index(">标准单位成本<"), html.index(">单位成本差异<"))

        row = next(
            r for r in resp.context["rows"] if r.work_order.order_no == "MO-STD-1"
        )
        # 标准材料单位成本 20 ×(1+20%+30%) = 30；实际单位成本 15；差异 = 15 − 30 = −15
        self.assertEqual(row.standard_unit_cost, Decimal("30"))
        self.assertEqual(row.unit_cost_diff, Decimal("-15"))
        self.assertEqual(row.unit_cost_diff_class, "woc-diff-good")
        # 用户 2026-09-24：单位成本差异改成显示百分比、保留两位小数（−15 ÷ 30 = −50.00%）
        self.assertEqual(row.unit_cost_diff_pct, Decimal("-50"))
        self.assertIn("-50.00%", html)
        # 合计行：按完工数量加权
        self.assertEqual(resp.context["subtotal"]["standard_unit_cost"], Decimal("30"))
        self.assertEqual(resp.context["subtotal"]["unit_cost_diff"], Decimal("-15"))
        self.assertEqual(
            resp.context["subtotal"]["unit_cost_diff_pct"], Decimal("-50")
        )

        # 导出 Excel 也带上这两列（与页面同口径：差异写百分比文本）
        exp = self.client.get(
            reverse("export_listing", args=[self.period.id, "work_order_cost"])
        )
        wb = load_workbook(BytesIO(exp.content))
        ws = wb.active
        header_row = [c.value for c in ws[1]]
        self.assertIn("标准单位成本", header_row)
        self.assertIn("单位成本差异", header_row)
        self.assertEqual(
            header_row.index("单位成本差异"), header_row.index("标准单位成本") + 1
        )
        self.assertEqual(
            header_row.index("标准单位成本"), header_row.index("完工单位成本") + 1
        )
        diff_col = header_row.index("单位成本差异")
        order_col = header_row.index("工单编号")
        detail = [
            r for r in ws.iter_rows(min_row=2, values_only=True)
            if r[order_col] == "MO-STD-1"
        ]
        self.assertEqual(len(detail), 1)
        self.assertEqual(detail[0][diff_col], "-50.00%")

    def test_standard_cost_page_has_no_actual_side(self):
        """用户 2026-09-22：标准成本页不再重复实际成本与差异（那在工单成本页）；
        用户 2026-09-23：占比表里也不要「本月实际（参考）」列。"""
        html = self.client.get(
            reverse("standard_costs", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertIn("标准材料单位成本", html)
        for gone in ("实际总成本", "差异额", "差异率", "实际制造费用", "本月实际"):
            self.assertNotIn(gone, html)
        # 指向工单成本页
        self.assertIn("工单成本", html)

    def test_nav_has_the_entry_after_work_order_cost(self):
        resp = self.client.get(reverse("standard_costs", args=[self.period.id]))
        nav = resp.content.decode("utf-8").split('data-nav="production"', 1)[1].split(
            "</div>", 2
        )[0]
        self.assertIn("标准成本", nav)
        self.assertIn(reverse("standard_costs", args=[self.period.id]), nav)
        self.assertLess(nav.index("工单成本"), nav.index("标准成本"))

    # ---------- 用户 2026-09-23 追加：「转换后子项用量:分子」列 + 标准成本页 BOM 明细表 ----------

    def test_bom_converted_numerator_formula(self):
        """两列折算口径（用户 2026-09-30）：

        * 「用量(转100%)」＝用量（分子÷分母）÷ 该工单 Σ用量；
        * 「转换后子项用量:分子」＝**标准用量 ÷ 该工单 Σ标准用量**（分子用标准用量，合计＝1＝100%）；
          该工单有行缺标准用量时退回用量口径。

        例：机织 39.034H1-HF20250734-3：标准用量 1450.17/25.83/82、合计 1558 →
        0.930789473684211 / 0.0165789473684211 / 0.0526315789473684（合计 100%）。
        """
        from settlement.views import (
            bom_converted_numerator,
            bom_normalized_qty,
            bom_qty_is_scaled,
            bom_share_resolver,
        )

        # 纯函数：一律折算（不再看 Σ用量 是否超过 100%）
        self.assertFalse(bom_qty_is_scaled(Decimal("0.6")))     # 未超 100%
        self.assertEqual(
            bom_normalized_qty(Decimal("0.1"), Decimal("0.6")),
            Decimal("0.1") / Decimal("0.6"),
        )
        self.assertEqual(
            bom_normalized_qty(Decimal("0.3537"), Decimal("0.38")).quantize(Decimal("0.000001")),
            Decimal("0.930789"),
        )
        self.assertTrue(bom_qty_is_scaled(Decimal("2")))
        # bom_converted_numerator 只剩「没有标准用量」时的兜底口径（用量 × 100）
        self.assertEqual(
            bom_converted_numerator(Decimal("0.5"), Decimal("1")), Decimal("50.0")
        )

        # 走 resolver（有标准用量的工单）：分子＝标准用量
        batch = ImportBatch.objects.create(
            period=self.period, kind=ImportBatch.Kind.BOM, original_name="bom-share"
        )
        # 用户举的工单：用量 0.3537/0.0063/0.02（Σ=0.38）、标准用量 1450.17/25.83/82（Σ=1558）
        for code, num, den, std in (
            ("RM-S1", "35.37", "100", "1450.17"),
            ("RM-S2", "0.63", "100", "25.83"),
            ("RM-S3", "2", "100", "82"),
        ):
            BomLine.objects.create(
                period=self.period, batch=batch,
                work_order_no="39.034H1-HF20250734-3", material_code="P-SHARE",
                sub_item_code=code, qty_numerator=Decimal(num),
                qty_denominator=Decimal(den), standard_qty=Decimal(std),
            )
        share = bom_share_resolver(self.period)
        rows = list(self.period.bom_lines.filter(work_order_no="39.034H1-HF20250734-3"))
        got = [tuple(Decimal(x).quantize(Decimal("0.000001")) for x in share(row)) for row in rows]
        self.assertEqual(
            got,
            [
                (Decimal("0.930789"), Decimal("0.930789")),
                (Decimal("0.016579"), Decimal("0.016579")),
                (Decimal("0.052632"), Decimal("0.052632")),
            ],
        )
        self.assertEqual(sum(r[1] for r in got), Decimal("1.000000"))

        # 两种口径不同时，以「标准用量」为准（用量 0.3333/0.6667 vs 标准用量 0.6667/1.3333）
        for code, num, std in (("RM-T1", "0.3333", "0.6667"), ("RM-T2", "0.6667", "1.3333")):
            BomLine.objects.create(
                period=self.period, batch=batch, work_order_no="MO-STD-BASIS",
                material_code="P-T", sub_item_code=code,
                qty_numerator=Decimal(num), qty_denominator=Decimal("1"),
                standard_qty=Decimal(std),
            )
        share = bom_share_resolver(self.period)
        row = self.period.bom_lines.get(sub_item_code="RM-T1")
        qty_ratio, converted = share(row)
        self.assertEqual(qty_ratio.quantize(Decimal("0.000001")), Decimal("0.333300"))   # 用量口径
        self.assertEqual(converted.quantize(Decimal("0.000001")), Decimal("0.333350"))   # 标准用量口径

    def test_bom_page_has_converted_numerator_column(self):
        """BOM 表台账页：新列加在「用量(转100%)」之后，值＝折算后用量×100。"""
        batch = ImportBatch.objects.create(
            period=self.period, kind=ImportBatch.Kind.BOM, original_name="bom-conv"
        )
        # 分母写法混用：97.86/1 + 2.14/1 → Σ用量 = 100（>100% 要折算到 100%）
        for code, num in (("RM-X", Decimal("97.86")), ("RM-Y", Decimal("2.14"))):
            BomLine.objects.create(
                period=self.period, batch=batch, work_order_no="MO-CONV",
                material_code="P-CONV", sub_item_code=code,
                qty_numerator=num, qty_denominator=Decimal("1"),
            )
        # 正常写法：10/100 + 50/100（Σ用量 = 0.6 → 不折算）
        for code, num in (("RM-A", Decimal("10")), ("RM-B", Decimal("50"))):
            BomLine.objects.create(
                period=self.period, batch=batch, work_order_no="MO-OK",
                material_code="P-OK", sub_item_code=code,
                qty_numerator=num, qty_denominator=Decimal("100"),
            )

        resp = self.client.get(reverse("bom_imports", args=[self.period.id]))
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        self.assertIn("转换后子项用量:分子", html)
        # 列序（只看表头区）：用量 → 用量(转100%) → 转换后子项用量:分子 → 单位
        head = html[html.index("<thead>"):html.index("</thead>")]
        self.assertLess(head.index(">用量(转100%)<"), head.index("转换后子项用量:分子"))
        self.assertLess(head.index("转换后子项用量:分子"), head.index(">单位<"))

        lines = {
            (line.work_order_no, line.sub_item_code): line
            for line in resp.context["page"].object_list
        }
        # 没有标准用量 → 退回用量口径：0.9786 / 0.0214（合计 1）
        self.assertEqual(
            lines[("MO-CONV", "RM-X")].qty_converted_numerator.quantize(Decimal("0.0001")),
            Decimal("0.9786"),
        )
        self.assertEqual(
            lines[("MO-CONV", "RM-Y")].qty_converted_numerator.quantize(Decimal("0.0001")),
            Decimal("0.0214"),
        )
        # 用户 2026-09-30：Σ用量 没超 100% 也一律折算（0.1/0.6、0.5/0.6 → 0.1667 / 0.8333）
        self.assertEqual(
            lines[("MO-OK", "RM-A")].qty_converted_numerator.quantize(Decimal("0.0001")),
            Decimal("0.1667"),
        )
        self.assertEqual(
            lines[("MO-OK", "RM-B")].qty_converted_numerator.quantize(Decimal("0.0001")),
            Decimal("0.8333"),
        )
        # 用量(转100%) 同口径：Σ用量 0.6 的工单也要折算
        self.assertEqual(
            lines[("MO-OK", "RM-A")].qty_normalized.quantize(Decimal("0.0001")),
            Decimal("0.1667"),
        )

    def test_bom_page_qty_total_uses_whole_work_order(self):
        """用户 2026-09-23：筛选取数只影响显示的行，Σ用量 仍按整张工单合计（否则折算失真）。"""
        batch = ImportBatch.objects.create(
            period=self.period, kind=ImportBatch.Kind.BOM, original_name="bom-conv2"
        )
        for code, num in (("RM-X", Decimal("97.86")), ("RM-Y", Decimal("2.14"))):
            BomLine.objects.create(
                period=self.period, batch=batch, work_order_no="MO-CONV",
                material_code="P-CONV", sub_item_code=code,
                qty_numerator=num, qty_denominator=Decimal("1"),
            )
        resp = self.client.get(
            reverse("bom_imports", args=[self.period.id]), {"sub_item_code": "RM-X"}
        )
        lines = list(resp.context["page"].object_list)
        self.assertEqual(len(lines), 1)
        self.assertEqual(lines[0].qty_total, Decimal("100.00"))     # 整张工单的 Σ用量
        self.assertTrue(lines[0].qty_is_scaled)
        # 没有标准用量 → 用量口径：97.86 ÷ 100 ＝ 0.9786
        self.assertEqual(
            lines[0].qty_converted_numerator.quantize(Decimal("0.0001")),
            Decimal("0.9786"),
        )

    def test_bom_export_has_converted_numerator_column(self):
        """BOM 表导出：与页面同列同口径。"""
        batch = ImportBatch.objects.create(
            period=self.period, kind=ImportBatch.Kind.BOM, original_name="bom-conv3"
        )
        for code, num in (("RM-X", Decimal("97.86")), ("RM-Y", Decimal("2.14"))):
            BomLine.objects.create(
                period=self.period, batch=batch, work_order_no="MO-CONV",
                material_code="P-CONV", sub_item_code=code,
                qty_numerator=num, qty_denominator=Decimal("1"),
            )
        resp = self.client.get(reverse("export_listing", args=[self.period.id, "bom"]))
        wb = load_workbook(BytesIO(resp.content), read_only=True, data_only=True)
        rows = list(wb[wb.sheetnames[0]].iter_rows(values_only=True))
        wb.close()
        header = list(rows[0])
        self.assertIn("转换后子项用量:分子", header)
        self.assertEqual(
            header.index("转换后子项用量:分子"), header.index("用量(转100%)") + 1
        )
        conv = header.index("转换后子项用量:分子")
        first = next(r for r in rows[1:] if r[0] == "MO-CONV")
        # 没有标准用量 → 退回用量口径：97.86 ÷ 100 ＝ 0.9786
        self.assertAlmostEqual(float(first[conv]), 0.9786, places=4)

    def test_bom_export_mirrors_uploaded_file_columns(self):
        """用户 2026-09-30：BOM 表导出＝上传的 BOM 文件列（原样同序）＋ 系统列追加在最右边。"""
        from openpyxl import Workbook

        header = [
            "序列", "序号", "工单编号", "物料编号", "物料名称", "项次",
            "子项物料编号", "子项物料名称", "子项用量:分子", "子项用量:分母",
            "单位名称", "损耗率(%)", "标准用量", "需求数",
        ]
        wb = Workbook()
        ws = wb.active
        ws.title = "BOM表"
        ws.append(header)
        # 同一工单两行（项次 10 / 20）＋ 另一张工单一行
        ws.append([1, 1, "MO-UP-1", "P-UP", "上传产品一", 10, "RM-U1", "上传原料一",
                   80, 100, "kg", 0, 0.8, 0.8])
        ws.append([2, 1, "MO-UP-1", "P-UP", "上传产品一", 20, "RM-U2", "上传原料二",
                   20, 100, "kg", 0, 0.2, 0.2])
        ws.append([3, 2, "MO-UP-2", "P-UP2", "上传产品二", 10, "RM-U3", "上传原料三",
                   10, 100, "kg", 0, 0.1, 0.1])
        stream = BytesIO()
        wb.save(stream)
        stream.seek(0)
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                batch = create_import_batch(
                    period=self.period,
                    kind=ImportBatch.Kind.BOM,
                    uploaded_file=SimpleUploadedFile(
                        "BOM表.xlsx", stream.read(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ),
                    user=self.user,
                )
                process_batch(batch)
                self.assertEqual(self.period.bom_lines.count(), 4)  # 上传 3 行 + setUp 的 1 行
                resp = self.client.get(
                    reverse("export_listing", args=[self.period.id, "bom"])
                )
                wb2 = load_workbook(BytesIO(resp.content), read_only=True, data_only=True)
                rows = list(wb2[wb2.sheetnames[0]].iter_rows(values_only=True))
                wb2.close()

        # 表头 = 上传表头原样（14 列） + 系统列（追加在最后）
        self.assertEqual(
            [str(c) for c in rows[0]],
            header + ["用量", "用量(转100%)", "转换后子项用量:分子", "单价", "标准成品单价"],
        )
        # 序列 ＝ 行号；序号 ＝ 同一工单一个号（换工单才 +1）
        for _i, _row in enumerate(rows[1:], start=1):
            self.assertEqual(int(_row[0]), _i)
        by_order = {}
        for _row in rows[1:]:
            by_order.setdefault(_row[2], []).append(int(_row[1]))
        self.assertEqual(by_order["MO-UP-1"], [by_order["MO-UP-1"][0]] * 2)
        self.assertEqual(by_order["MO-UP-2"][0], by_order["MO-UP-1"][0] + 1)
        # 上传列的值逐列映射回来（数据＝库里的记录，列＝上传文件的列）
        row = next(r for r in rows[1:] if r[6] == "RM-U1")
        self.assertEqual(row[2], "MO-UP-1")
        self.assertEqual(row[3], "P-UP")
        self.assertEqual(row[4], "上传产品一")
        self.assertEqual(str(row[5]), "10")          # 项次
        self.assertEqual(row[7], "上传原料一")
        self.assertEqual(float(row[8]), 80.0)        # 子项用量:分子
        self.assertEqual(float(row[9]), 100.0)       # 子项用量:分母
        self.assertEqual(row[10], "kg")
        self.assertEqual(float(row[12]), 0.8)        # 标准用量（上传原值）
        # 系统列：用量 ＝ 80/100；Σ标准用量 ＝ 1.0 → 转换后 ＝ 标准用量 ÷ Σ标准用量 ＝ 0.8
        self.assertEqual(float(row[14]), 0.8)
        self.assertEqual(float(row[15]), 0.8)
        self.assertEqual(float(row[16]), 0.8)
        self.assertIsNone(row[17])                   # 单价（结算后才回填）

    def test_bom_export_keeps_uploaded_price_columns(self):
        """上传文件自带「单价（子项物料的单价）」「标准成品单价」时，系统不重复追加同义列。"""
        from openpyxl import Workbook

        header = [
            "序号", "工单编号", "物料编号", "物料名称", "项次", "子项物料编号",
            "子项物料名称", "子项用量:分子", "子项用量:分母", "单位名称",
            "单价（子项物料的单价）", "标准成品单价",
        ]
        wb = Workbook()
        ws = wb.active
        ws.title = "BOM表"
        ws.append(header)
        ws.append([1, "MO-OLD-1", "P-OLD", "老格式产品", 10, "RM-O1", "老格式原料",
                   10, 100, "kg", 12.5, 20])
        stream = BytesIO()
        wb.save(stream)
        stream.seek(0)
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                batch = create_import_batch(
                    period=self.period,
                    kind=ImportBatch.Kind.BOM,
                    uploaded_file=SimpleUploadedFile(
                        "BOM表.xlsx", stream.read(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ),
                    user=self.user,
                )
                process_batch(batch)
                resp = self.client.get(
                    reverse("export_listing", args=[self.period.id, "bom"])
                )
                wb2 = load_workbook(BytesIO(resp.content), read_only=True, data_only=True)
                rows = list(wb2[wb2.sheetnames[0]].iter_rows(values_only=True))
                wb2.close()

        # 单价 / 标准成品单价 上传列里已有 → 只追加另外三个系统列
        self.assertEqual(
            [str(c) for c in rows[0]],
            header + ["用量", "用量(转100%)", "转换后子项用量:分子"],
        )
        row = next(r for r in rows[1:] if r[5] == "RM-O1")
        self.assertEqual(float(row[10]), 12.5)       # 单价（子项物料的单价）
        self.assertEqual(float(row[11]), 20.0)       # 标准成品单价
        self.assertEqual(float(row[12]), 0.1)        # 用量 ＝ 10/100
        # 这张工单只有一行 → 占比 ＝ 1（100%）；该上传没有标准用量列 → 走用量口径
        self.assertEqual(float(row[13]), 1.0)        # 用量(转100%) ＝ 0.1 ÷ 0.1
        self.assertEqual(float(row[14]), 1.0)        # 转换后子项用量:分子 ＝ 1（100%）

    def test_standard_cost_page_has_bom_detail_table(self):
        """用户 2026-09-23：标准成本页加「BOM 明细表」，说明标准材料单位成本的由来。"""
        resp = self.client.get(reverse("standard_costs", args=[self.period.id]))
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode("utf-8")
        self.assertIn("<h2>BOM 明细表</h2>", html)
        self.assertLess(
            html.index("<h2>工单标准成本计算表</h2>"), html.index("<h2>BOM 明细表</h2>")
        )
        # 新列也在（与 BOM 表台账页同一列）
        self.assertIn("转换后子项用量:分子", html)
        rows = resp.context["bom_rows"]
        self.assertEqual(len(rows), 1)          # setUp 里 MO-STD-1 只有一行 BOM
        row = rows[0]
        self.assertEqual(row["order_no"], "MO-STD-1")
        self.assertEqual(row["workshop_label"], "染色车间")
        self.assertEqual(row["product_code"], "P-STD")
        self.assertEqual(row["sub_item_code"], "RM-STD")
        self.assertEqual(row["qty"], Decimal("2"))            # 2 ÷ 1
        self.assertEqual(row["amount"], Decimal("20"))        # 2 × 10 ＝ 标准成品单价
        self.assertEqual(row["standard_price"], Decimal("20"))
        # Σ用量 = 2 > 1 → 该行占整单 100%（没有标准用量 → 用量口径：2/2 ＝ 1）
        self.assertTrue(row["qty_is_scaled"])
        self.assertEqual(row["qty_converted_numerator"], Decimal("1"))

    def test_standard_cost_bom_detail_follows_filters(self):
        """BOM 明细表按自己的期间/车间/工单号字段筛选（用户 2026-09-24）。"""
        url = reverse("standard_costs", args=[self.period.id])
        self.assertEqual(self.client.get(url).context["bom_row_count"], 1)
        # 工单号模糊匹配
        self.assertEqual(
            self.client.get(url, {"b_order_no": "MO-STD"}).context["bom_row_count"], 1
        )
        self.assertEqual(
            self.client.get(url, {"b_order_no": "NOPE"}).context["bom_row_count"], 0
        )
        # 车间
        self.assertEqual(
            self.client.get(url, {"b_workshop": "RanZheng"}).context["bom_row_count"], 1
        )
        self.assertEqual(
            self.client.get(url, {"b_workshop": "tefang"}).context["bom_row_count"], 0
        )
        # 期间：本期有 BOM 明细；换到没有结算结果的期间（8 月）就为空
        AccountingPeriod.objects.create(month=date(2026, 8, 1))
        self.assertEqual(
            self.client.get(url, {"b_period": "2026-08"}).context["bom_row_count"], 0
        )
        self.assertEqual(
            self.client.get(url, {"b_period": "2026-07"}).context["bom_row_count"], 1
        )

    def test_bom_detail_section_has_its_own_filter_row(self):
        """用户 2026-09-24：BOM 明细表加期间/车间/工单编号筛选字段 + 导出Excel/打印本表，
        并且与筛选字段同一行、底对齐。"""
        html = self.client.get(
            reverse("standard_costs", args=[self.period.id])
        ).content.decode("utf-8")
        # 第三张表那一节的筛选行（前面两张表各有自己的 sc-filter-row）
        start = html.index("<h2>BOM 明细表</h2>")
        block = html[start:]
        form = block[: block.index("</form>")]
        self.assertIn('name="b_period"', form)
        self.assertIn('name="b_workshop"', form)
        self.assertIn('name="b_order_no"', form)
        actions = form[form.index('<div class="filter-actions">'):]
        for label in ("筛选", "清空", "导出Excel"):
            self.assertIn(label, actions)
        self.assertIn('class="print-here"', actions)
        # 四个动作按 筛选 → 清空 → 导出Excel → 打印本表 排在同一行（对齐）
        self.assertLess(actions.index("筛选"), actions.index("清空"))
        self.assertLess(actions.index("清空"), actions.index("导出Excel"))
        self.assertLess(actions.index("导出Excel"), actions.index("print-here"))
        # 筛选行用的是与另外两节同款的对齐样式（flex-wrap + 底对齐）
        self.assertIn('<form class="sc-filter-row" method="get">', form)
        # 导出链接带本表的筛选参数（筛了车间/工单号时）
        filtered = self.client.get(
            reverse("standard_costs", args=[self.period.id]),
            {"b_workshop": "RanZheng", "b_order_no": "MO-STD"},
        ).content.decode("utf-8")
        filtered = filtered[filtered.index("<h2>BOM 明细表</h2>"):]
        filtered = filtered[: filtered.index("</form>")]
        self.assertIn("standard_cost_bom", filtered)
        self.assertIn("b_workshop=RanZheng", filtered)
        self.assertIn("b_order_no=MO-STD", filtered)

    def test_bom_detail_qty_columns_keep_two_decimals(self):
        """用户 2026-09-24：分子/分母/用量/用量(转100%)/转换后子项用量:分子 保留两位小数，
        并且「用量(转100%)」不再带「（Σ用量 … → 折算）」标注。"""
        batch = ImportBatch.objects.create(
            period=self.period, kind=ImportBatch.Kind.BOM, original_name="bom-dec"
        )
        BomLine.objects.create(
            period=self.period, batch=batch, work_order_no="MO-STD-1",
            material_code="P-STD", sub_item_code="RM-DEC",
            qty_numerator=Decimal("97.86"), qty_denominator=Decimal("1"),
            unit_price=Decimal("10"), standard_price=Decimal("20"),
        )
        html = self.client.get(
            reverse("standard_costs", args=[self.period.id])
        ).content.decode("utf-8")
        block = html[html.index("<h2>BOM 明细表</h2>"):]
        # 数量列两位小数（RM-DEC 那行：97.86 / 1.00 / 97.86）
        self.assertIn(">97.86<", block)
        self.assertIn(">1.00<", block)
        # 原来的六位小数与折算标注都不见了
        self.assertNotIn("97.860000", block)
        # 单元格里的「（Σ用量 2.00 → 折算）」小字标注已删除（表头 tooltip 里的说明保留）
        self.assertNotIn("→ 折算）", block)
        self.assertNotIn('<span style="color:#8a97a3;font-size:11px;">', block)

    def test_bom_detail_lists_bom_lines_without_cost_result(self):
        """用户 2026-09-30：BOM 明细表列**本期全部 BOM 行**——没有结算结果的工单也要列出来。"""
        batch = ImportBatch.objects.create(
            period=self.period, kind=ImportBatch.Kind.BOM, original_name="bom-all"
        )
        WorkOrder.objects.create(
            period=self.period, batch=batch, order_no="MO-NORESULT",
            production_dept_code="tefang", plan_quantity=Decimal("5"),
        )
        BomLine.objects.create(
            period=self.period, batch=batch, work_order_no="MO-NORESULT",
            material_code="P-NR", material_name="无结算产品",
            sub_item_code="RM-NR", qty_numerator=Decimal("1"),
            qty_denominator=Decimal("10"),
        )
        resp = self.client.get(reverse("standard_costs", args=[self.period.id]))
        # setUp 的 MO-STD-1（有结算）+ 本次的 MO-NORESULT（没结算）＝ 2 行 2 张工单
        self.assertEqual(resp.context["bom_row_count"], 2)
        self.assertEqual(resp.context["bom_order_count"], 2)
        rows = {r["order_no"]: r for r in resp.context["bom_rows"]}
        self.assertIn("MO-NORESULT", rows)
        row = rows["MO-NORESULT"]
        # 车间取工单主档；产品编码/名称用 BOM 成品物料兜底
        self.assertEqual(row["workshop_label"], "纺纱车间")
        self.assertEqual(row["product_code"], "P-NR")
        self.assertEqual(row["product_name"], "无结算产品")
        self.assertIsNone(row["standard_price"])   # 没结算 → 空
        html = resp.content.decode("utf-8")
        self.assertIn("本表共 2 行、2 张工单", html)

    def test_bom_detail_columns_mirror_uploaded_file(self):
        """用户 2026-09-30：BOM 明细表及其导出＝上传文件的列原样 + 系统列追加在最右边。"""
        from openpyxl import Workbook

        header = [
            "序列", "序号", "工单编号", "物料编号", "物料名称", "项次",
            "子项物料编号", "子项物料名称", "子项用量:分子", "子项用量:分母", "单位名称",
        ]
        wb = Workbook()
        ws = wb.active
        ws.title = "BOM表"
        ws.append(header)
        ws.append([1, 1, "MO-MIR", "P-MIR", "镜像产品", 10, "RM-M1", "镜像原料一",
                   60, 100, "kg"])
        ws.append([2, 1, "MO-MIR", "P-MIR", "镜像产品", 20, "RM-M2", "镜像原料二",
                   40, 100, "kg"])
        stream = BytesIO()
        wb.save(stream)
        stream.seek(0)
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                batch = create_import_batch(
                    period=self.period,
                    kind=ImportBatch.Kind.BOM,
                    uploaded_file=SimpleUploadedFile(
                        "BOM表.xlsx", stream.read(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ),
                    user=self.user,
                )
                process_batch(batch)
                page = self.client.get(reverse("standard_costs", args=[self.period.id]))
                export = self.client.get(
                    reverse("export_listing", args=[self.period.id, "standard_cost_bom"])
                )
                wb2 = load_workbook(BytesIO(export.content), read_only=True, data_only=True)
                rows = [list(r) for r in wb2[wb2.sheetnames[0]].iter_rows(values_only=True)]
                wb2.close()

        expected = header + [
            "用量", "用量(转100%)", "转换后子项用量:分子", "单价", "金额", "标准成品单价",
        ]
        # 导出行列＝上传列原样 + 系统列（在最右边）
        self.assertEqual([str(c) for c in rows[0]], expected)
        mir = [r for r in rows[1:] if r[2] == "MO-MIR"]
        self.assertEqual(len(mir), 2)
        # 行序＝上传顺序：setUp 里那行在前（序列 1、序号 1），上传的两行在后（序列 2、3，序号 2）
        self.assertEqual([int(r[0]) for r in mir], [2, 3])   # 序列＝行号
        self.assertEqual([int(r[1]) for r in mir], [2, 2])   # 序号＝同一工单一个号
        self.assertEqual(mir[0][6], "RM-M1")
        self.assertEqual(float(mir[0][8]), 60.0)
        self.assertEqual(float(mir[0][9]), 100.0)
        self.assertEqual(mir[0][10], "kg")
        # 系统列：用量 ＝ 60/100；Σ用量 ＝ 1.0 → 用量(转100%) 同值；该上传没标准用量 → 转换后同口径 0.6
        self.assertEqual(float(mir[0][11]), 0.6)
        self.assertEqual(float(mir[0][12]), 0.6)
        self.assertEqual(float(mir[0][13]), 0.6)
        self.assertIsNone(mir[0][14])                        # 单价（结算后才回填）
        # 页面表头同列同序（在最右边）
        html = page.content.decode("utf-8")
        head = html[html.index("<h2>BOM 明细表</h2>"):]
        head = head[: head.index("</thead>")]
        for name in expected:
            self.assertIn(f">{name}<", head)
        self.assertNotIn(">产品编码<", head)                  # 列里不再插系统列（期间/车间/产品…）
        self.assertIn("本表共 3 行、2 张工单", html)          # setUp 1 行 + 上传 2 行

    def test_standard_cost_export_has_bom_detail_sheet(self):
        """标准成本导出：第三张表「BOM 明细」，列与页面一致。"""
        resp = self.client.get(
            reverse("export_listing", args=[self.period.id, "standard_cost"])
        )
        wb = load_workbook(BytesIO(resp.content))
        self.assertEqual(
            wb.sheetnames,
            ["工单标准成本计算表", "标准成本构成表", "BOM 明细"],
        )
        sheet = wb["BOM 明细"]
        header = [cell.value for cell in sheet[1]]
        self.assertEqual(
            header.index("转换后子项用量:分子"), header.index("用量(转100%)") + 1
        )
        values = [cell.value for cell in sheet[2]]
        self.assertEqual(values[0:5], ["2026-07", "染色车间", "MO-STD-1", "P-STD", "标准成本测试产品"])
        self.assertEqual(values[6], "RM-STD")

    def test_standard_cost_bom_export_is_standalone(self):
        """用户 2026-09-24：BOM 明细表可以单独导出 Excel（一张表，列与页面一致），
        并按本表自己的车间/工单号筛选取数。"""
        url = reverse("export_listing", args=[self.period.id, "standard_cost_bom"])
        resp = self.client.get(url)
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(BytesIO(resp.content))
        self.assertEqual(wb.sheetnames, ["BOM 明细"])
        sheet = wb["BOM 明细"]
        header = [cell.value for cell in sheet[1]]
        self.assertEqual(header[0], "期间")
        self.assertEqual(header[8], "子项用量:分子")
        self.assertEqual(header[9], "子项用量:分母")
        self.assertEqual(header[10], "用量")
        self.assertEqual(header[11], "用量(转100%)")
        self.assertEqual(header[12], "转换后子项用量:分子")
        self.assertEqual(header[16], "标准成品单价")
        values = [cell.value for cell in sheet[2]]
        self.assertEqual(values[0:5], ["2026-07", "染色车间", "MO-STD-1", "P-STD", "标准成本测试产品"])
        # 筛不到 → 只有表头
        resp = self.client.get(url, {"b_order_no": "NOPE"})
        wb = load_workbook(BytesIO(resp.content))
        self.assertEqual(wb["BOM 明细"].max_row, 1)
        # 导出中心里也有这一项
        resp = self.client.get(reverse("export_center", args=[self.period.id]))
        self.assertContains(resp, "标准成本BOM明细")


class BackfillBomColumnsTests(SettlementFixtureMixin, TestCase):
    """用户 2026-09-30：`backfill_bom_columns` —— 用当初上传的 BOM 文件补齐
    「损耗率(%) / 标准用量 / 需求数」三列（这三列是后来才加到平台的，老批次全是 NULL）。"""

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="bom-fill", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))

    def _upload(self, data_rows):
        from openpyxl import Workbook

        wb = Workbook()
        ws = wb.active
        ws.title = "BOM表"
        ws.append([
            "序列", "序号", "工单编号", "物料编号", "物料名称", "项次",
            "子项物料编号", "子项物料名称", "子项用量:分子", "子项用量:分母",
            "单位名称", "损耗率(%)", "标准用量", "需求数",
        ])
        for row in data_rows:
            ws.append(row)
        stream = BytesIO()
        wb.save(stream)
        stream.seek(0)
        batch = create_import_batch(
            period=self.period,
            kind=ImportBatch.Kind.BOM,
            uploaded_file=SimpleUploadedFile(
                "BOM表.xlsx", stream.read(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ),
            user=self.user,
        )
        process_batch(batch)
        return batch

    def test_backfills_three_columns_from_uploaded_file(self):
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                self._upload([
                    [1, 1, "MO-F1", "P-F1", "产品一", 10, "RM-F1", "原料一",
                     60, 100, "kg", 5, 0.6, 0.63],
                    [2, 1, "MO-F1", "P-F1", "产品一", 20, "RM-F2", "原料二",
                     40, 100, "kg", 0, 0.4, 0.4],
                ])
                # 新导入本来就带这三列
                self.assertEqual(
                    self.period.bom_lines.get(sub_item_code="RM-F1").standard_qty,
                    Decimal("0.600000"),
                )
                # 模拟「当年导入时平台还没有这三列」：全部清空
                self.period.bom_lines.update(
                    loss_rate=None, standard_qty=None, required_qty=None
                )
                out = StringIO()
                call_command("backfill_bom_columns", "--dry-run", stdout=out)
                self.assertIsNone(
                    self.period.bom_lines.get(sub_item_code="RM-F1").loss_rate
                )  # dry-run 不写库

                call_command("backfill_bom_columns", stdout=out)
                line = self.period.bom_lines.get(sub_item_code="RM-F1")
                self.assertEqual(line.loss_rate, Decimal("5.0000"))
                self.assertEqual(line.standard_qty, Decimal("0.600000"))
                self.assertEqual(line.required_qty, Decimal("0.630000"))
                self.assertEqual(
                    self.period.bom_lines.filter(standard_qty__isnull=True).count(), 0
                )

                # 幂等：再跑一次值不变；且**不覆盖**已有值（手工改过的值要留着）
                self.period.bom_lines.filter(sub_item_code="RM-F2").update(
                    standard_qty=Decimal("9.99")
                )
                call_command("backfill_bom_columns", stdout=out)
                self.assertEqual(
                    self.period.bom_lines.get(sub_item_code="RM-F2").standard_qty,
                    Decimal("9.99"),
                )
                self.assertEqual(
                    self.period.bom_lines.get(sub_item_code="RM-F2").required_qty,
                    Decimal("0.400000"),
                )


class FinishProgressColumnTests(SettlementFixtureMixin, TestCase):
    """用户 2026-09-24：「本期完工 / 累计完工 / 本期完工率 / 累计完工率」口径改版。

    旧口径（已改）：本期完工 ＝ 累计完工 ＝ 本月报工，完工率 ＝ 完工÷(完工＋在产品)
      → 两对列完全重复（同一张表里「报工数量」也是这个数，三列同一个值）；
        完工率其实是成本拆分比例，出现「计划 9,280 / 完工 32.8 显示 100.00%」。
    新口径：本期完工＝本月报工；累计完工＝工单明细的「入库数量」（ERP 跨月累计，
            无明细退回本月报工）；两个完工率＝完工 ÷ 计划数量 × 100（超产按实际显示）。

    见 `settlement/services/workorder_costing.py::attach_finish_progress`（页面/导出/结算工作簿共用）。
    """

    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="finish-progress", password="test-password"
        )
        self.client.force_login(self.user)
        self.period = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        self.batch = self.create_batch(
            self.period, ImportBatch.Kind.WORK_ORDER, "finish-progress"
        )
        detail_batch = self.create_batch(
            self.period, ImportBatch.Kind.WORK_ORDER_DETAIL, "finish-progress-detail"
        )
        # 跨月工单：计划 1,000、本月报工 200、ERP 累计入库 900、在产品 700
        self.order = WorkOrder.objects.create(
            period=self.period, batch=self.batch, order_no="MO-FIN-1",
            production_dept_code="RanZheng", plan_quantity=Decimal("1000"),
        )
        WorkOrderCostResult.objects.create(
            period=self.period, work_order=self.order, product_code="P-FIN",
            plan_quantity=Decimal("1000"), finished_quantity=Decimal("200"),
            wip_quantity=Decimal("700"), direct_material=Decimal("1000"),
            finished_cost=Decimal("300"), wip_cost=Decimal("700"),
            total_cost=Decimal("1000"),
        )
        WorkOrderDetail.objects.create(
            period=self.period, batch=detail_batch, order_no="MO-FIN-1",
            plan_quantity=Decimal("1000"), issue_quantity=Decimal("900"),
            receipt_quantity=Decimal("900"), completion_status="超时完成",
        )

    def _row(self, order_no):
        return next(
            r for r in self.client.get(
                reverse("work_order_costs", args=[self.period.id])
            ).context["rows"] if r.work_order.order_no == order_no
        )

    def test_accum_finished_comes_from_work_order_detail(self):
        row = self._row("MO-FIN-1")
        # 本期完工＝本月报工 200；累计完工＝工单明细入库数量 900（不再等于本期）
        self.assertEqual(row.current_finished, Decimal("200"))
        self.assertEqual(row.accum_finished, Decimal("900"))
        self.assertNotEqual(row.current_finished, row.accum_finished)
        # 完工率＝完工÷计划：200/1000＝20%、900/1000＝90%
        # （不再是 完工÷(完工＋在产品)＝200/(200+700)＝22.22%）
        self.assertEqual(row.current_finish_rate_pct, Decimal("20"))
        self.assertEqual(row.finish_rate_pct, Decimal("90"))

    def test_accum_falls_back_to_current_without_detail(self):
        """没有工单明细的工单：累计完工退回本月报工（不会变成 0 或空）。"""
        WorkOrderDetail.objects.filter(order_no="MO-FIN-1").delete()
        row = self._row("MO-FIN-1")
        self.assertEqual(row.accum_finished, Decimal("200"))
        self.assertEqual(row.finish_rate_pct, Decimal("20"))

    def test_rate_is_against_plan_not_wip(self):
        """回归：在产品被清零（视同完工／纺纱不用粗纱／机织重量法）时完工率不能再跳到 100%。

        用户 2026-09-24 实测的真实案例：计划 9,280 / 完工 32.8 / 在产品 0，
        旧口径显示 100.00%，新口径必须是 0.35%。
        """
        order = WorkOrder.objects.create(
            period=self.period, batch=self.batch, order_no="MO-FIN-SMALL",
            production_dept_code="tefang", plan_quantity=Decimal("9280"),
        )
        WorkOrderCostResult.objects.create(
            period=self.period, work_order=order, product_code="P-SMALL",
            plan_quantity=Decimal("9280"), finished_quantity=Decimal("32.8"),
            wip_quantity=Decimal("0"), direct_material=Decimal("100"),
            finished_cost=Decimal("100"), wip_cost=Decimal("0"),
            total_cost=Decimal("100"),
        )
        row = self._row("MO-FIN-SMALL")
        self.assertEqual(
            row.finish_rate_pct.quantize(Decimal("0.01")), Decimal("0.35")
        )
        self.assertEqual(
            row.current_finish_rate_pct.quantize(Decimal("0.01")), Decimal("0.35")
        )

    def test_over_production_is_shown_as_is(self):
        """超产（完工＞计划）原样显示，不封顶。"""
        order = WorkOrder.objects.create(
            period=self.period, batch=self.batch, order_no="MO-FIN-OVER",
            production_dept_code="BaoFu", plan_quantity=Decimal("10"),
        )
        WorkOrderCostResult.objects.create(
            period=self.period, work_order=order, product_code="P-OVER",
            plan_quantity=Decimal("10"), finished_quantity=Decimal("16.9"),
            wip_quantity=Decimal("0"), direct_material=Decimal("100"),
            finished_cost=Decimal("100"), wip_cost=Decimal("0"),
            total_cost=Decimal("100"),
        )
        row = self._row("MO-FIN-OVER")
        self.assertEqual(row.finish_rate_pct, Decimal("169.0"))

    def test_plan_zero_leaves_rates_blank(self):
        order = WorkOrder.objects.create(
            period=self.period, batch=self.batch, order_no="MO-FIN-NOPLAN",
            production_dept_code="tefang", plan_quantity=Decimal("0"),
        )
        WorkOrderCostResult.objects.create(
            period=self.period, work_order=order, product_code="P-NOPLAN",
            plan_quantity=Decimal("0"), finished_quantity=Decimal("5"),
            wip_quantity=Decimal("0"), direct_material=Decimal("10"),
            finished_cost=Decimal("10"), wip_cost=Decimal("0"),
            total_cost=Decimal("10"),
        )
        row = self._row("MO-FIN-NOPLAN")
        self.assertIsNone(row.current_finish_rate_pct)
        self.assertIsNone(row.finish_rate_pct)

    def test_page_and_export_show_the_same_numbers(self):
        html = self.client.get(
            reverse("work_order_costs", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertIn("20.00%", html)      # 本期完工率
        self.assertIn("90.00%", html)      # 累计完工率
        exp = self.client.get(
            reverse("export_listing", args=[self.period.id, "work_order_cost"])
        )
        wb = load_workbook(BytesIO(exp.content))
        sheet = wb.active
        header = [c.value for c in sheet[1]]
        row = next(
            r for r in sheet.iter_rows(min_row=2, values_only=True)
            if r[header.index("工单编号")] == "MO-FIN-1"
        )
        self.assertEqual(Decimal(str(row[header.index("本期完工")])), Decimal("200"))
        self.assertEqual(Decimal(str(row[header.index("累计完工")])), Decimal("900"))
        self.assertEqual(row[header.index("本期完工率")], "20.00%")
        self.assertEqual(row[header.index("累计完工率")], "90.00%")

    def test_settlement_workbook_uses_the_same_numbers(self):
        """结算工作簿的「工单成本」工作表同一口径（累计完工率＝900÷1000＝90%）。"""
        from settlement.services.exports import build_settlement_workbook

        wb = load_workbook(build_settlement_workbook(self.period))
        sheet = wb["工单成本"]
        header = [c.value for c in sheet[1]]
        row = next(
            r for r in sheet.iter_rows(min_row=2, values_only=True)
            if r[header.index("工单编号")] == "MO-FIN-1"
        )
        self.assertEqual(Decimal(str(row[header.index("累计完工")])), Decimal("900"))
        self.assertEqual(row[header.index("累计完工率")], "90.00%")
