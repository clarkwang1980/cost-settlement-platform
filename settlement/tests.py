import re

from datetime import date
from decimal import Decimal
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory

from django.conf import settings
from django.contrib.auth import get_user_model
from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import TestCase, override_settings
from django.urls import reverse
from openpyxl import load_workbook

from settlement.models import (
    AccountingPeriod,
    EndingInventoryBalance,
    HistoricalPrice,
    ImportBatch,
    InventoryTransaction,
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
    WorkOrder,
    WorkOrderCostResult,
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
from settlement.services.matching import auto_match_all
from settlement.services.purchase_voucher import build_purchase_voucher
from settlement.services.voucher_center import build_voucher_center
from settlement.services.purchase_accruals import (
    build_purchase_accrual_rows,
    build_purchase_difference_rows,
)
from settlement.services.workorder_costing import (
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
    build_sales_uninvoiced_rows,
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
        self.assertEqual(sales_issue.invoice_period, date(2026, 6, 1))

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

    def test_cross_period_sales_invoice_partially_clears_goods_issued(self):
        june = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        july = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        june_inventory_batch = self.create_batch(
            june,
            ImportBatch.Kind.INVENTORY,
            "sales-cross-inventory",
        )
        july_invoice_batch = self.create_batch(
            july,
            ImportBatch.Kind.SALES_INVOICE,
            "sales-cross-invoice",
        )
        product = Material.objects.create(
            code="FG-SALES-CROSS",
            name="跨月销售产品",
            uom="件",
        )
        sales_issue = self.create_txn(
            june,
            june_inventory_batch,
            txn_date=date(2026, 6, 25),
            document_no="SO-SALES-CROSS",
            business_type="销售出库",
            direction=InventoryTransaction.Direction.OUT,
            material=product,
            quantity=Decimal("10"),
            order_no="SALE-CROSS",
            customer_code="CUS-CROSS",
            customer_name="跨月客户",
        )
        sales_issue.unit_price = Decimal("100")
        sales_issue.amount = Decimal("1000")
        sales_issue.save(update_fields=["unit_price", "amount"])

        auto_match_all(june)
        june_rows = build_sales_uninvoiced_rows(june)
        self.assertEqual(len(june_rows), 1)
        self.assertEqual(june_rows[0].outstanding_quantity, Decimal("10"))
        self.assertEqual(june_rows[0].outstanding_cost, Decimal("1000.00"))
        june.status = AccountingPeriod.Status.LOCKED
        june.save(update_fields=["status"])

        invoice = SalesInvoiceLine.objects.create(
            period=july,
            batch=july_invoice_batch,
            invoice_no="SI-SALES-CROSS-07",
            invoice_date=date(2026, 7, 20),
            invoice_period=date(2026, 7, 1),
            sales_order_no="SALE-CROSS",
            customer_code="CUS-CROSS",
            customer_name="跨月客户",
            material_code=product.code,
            material_name=product.name,
            quantity=Decimal("4"),
            remaining_quantity=Decimal("4"),
        )

        auto_match_all(july)
        sales_issue.refresh_from_db()
        invoice.refresh_from_db()
        self.assertEqual(invoice.remaining_quantity, Decimal("0"))
        self.assertEqual(
            sales_issue.match_status,
            InventoryTransaction.MatchStatus.PARTIAL,
        )
        self.assertEqual(sales_issue.invoice_period, date(2026, 7, 1))

        july_rows = build_sales_uninvoiced_rows(july)
        self.assertEqual(len(july_rows), 1)
        self.assertTrue(july_rows[0].cross_period)
        self.assertEqual(july_rows[0].invoiced_quantity, Decimal("4"))
        self.assertEqual(july_rows[0].outstanding_quantity, Decimal("6"))
        self.assertEqual(july_rows[0].invoiced_cost, Decimal("400.00"))
        self.assertEqual(july_rows[0].outstanding_cost, Decimal("600.00"))
        self.assertEqual(july_rows[0].status, "部分开票")

        june_as_of_rows = build_sales_uninvoiced_rows(june)
        self.assertEqual(june_as_of_rows[0].outstanding_quantity, Decimal("10"))
        self.assertEqual(june_as_of_rows[0].outstanding_cost, Decimal("1000.00"))

        rollforward = build_goods_issued_rollforward(july)
        self.assertEqual(rollforward["opening_amount"], Decimal("1000.00"))
        self.assertEqual(rollforward["current_issue_amount"], Decimal("0.00"))
        self.assertEqual(
            rollforward["invoice_cleared_amount"],
            Decimal("400.00"),
        )
        self.assertEqual(rollforward["ending_amount"], Decimal("600.00"))
        self.assertEqual(rollforward["difference"], Decimal("0.00"))

        self.client.force_login(self.user)
        matches_response = self.client.get(
            reverse("sales_matches", kwargs={"period_id": july.id})
        )
        self.assertContains(matches_response, "SO-SALES-CROSS")
        self.assertContains(matches_response, "2026-06")
        uninvoiced_response = self.client.get(
            reverse("sales_uninvoiced", kwargs={"period_id": july.id})
        )
        self.assertContains(uninvoiced_response, "部分开票")
        self.assertContains(uninvoiced_response, "600.00")

    def test_incomplete_sales_keys_require_manual_cross_period_match(self):
        june = AccountingPeriod.objects.create(month=date(2026, 6, 1))
        july = AccountingPeriod.objects.create(month=date(2026, 7, 1))
        inventory_batch = self.create_batch(
            june,
            ImportBatch.Kind.INVENTORY,
            "sales-manual-inventory",
        )
        invoice_batch = self.create_batch(
            july,
            ImportBatch.Kind.SALES_INVOICE,
            "sales-manual-invoice",
        )
        product = Material.objects.create(
            code="FG-SALES-MANUAL",
            name="手工跨月产品",
            uom="件",
        )
        sales_issue = self.create_txn(
            june,
            inventory_batch,
            txn_date=date(2026, 6, 26),
            document_no="SO-SALES-MANUAL",
            business_type="销售出库",
            direction=InventoryTransaction.Direction.OUT,
            material=product,
            quantity=Decimal("3"),
            order_no="SALE-MANUAL",
            customer_code="CUS-MANUAL",
        )
        sales_issue.amount = Decimal("300")
        sales_issue.save(update_fields=["amount"])
        june.status = AccountingPeriod.Status.LOCKED
        june.save(update_fields=["status"])
        invoice = SalesInvoiceLine.objects.create(
            period=july,
            batch=invoice_batch,
            invoice_no="SI-SALES-MANUAL-07",
            invoice_date=date(2026, 7, 21),
            invoice_period=date(2026, 7, 1),
            sales_order_no="",
            customer_code="CUS-MANUAL",
            material_code=product.code,
            material_name=product.name,
            quantity=Decimal("3"),
            remaining_quantity=Decimal("3"),
        )

        auto_match_all(july)
        invoice.refresh_from_db()
        self.assertEqual(invoice.remaining_quantity, Decimal("3"))
        self.assertFalse(invoice.matchallocation_set.exists())

        self.client.force_login(self.user)
        detail_url = reverse(
            "source_document_detail",
            kwargs={
                "period_id": july.id,
                "source_type": "sales_invoice",
                "source_id": invoice.id,
            },
        )
        candidates_response = self.client.get(detail_url)
        self.assertContains(candidates_response, "SO-SALES-MANUAL")
        self.assertContains(candidates_response, "2026-06")
        response = self.client.post(
            detail_url,
            {
                "action": "add_match",
                "transaction_id": sales_issue.id,
                "matched_quantity": "3",
            },
        )
        self.assertEqual(response.status_code, 302)
        invoice.refresh_from_db()
        sales_issue.refresh_from_db()
        allocation = invoice.matchallocation_set.get()
        self.assertTrue(allocation.is_manual)
        self.assertEqual(invoice.match_mode, MatchingMode.MANUAL)
        self.assertEqual(invoice.remaining_quantity, Decimal("0"))
        self.assertEqual(
            sales_issue.match_status,
            InventoryTransaction.MatchStatus.MATCHED,
        )
        self.assertEqual(build_sales_uninvoiced_rows(july), [])

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
        difference_response = self.client.get(
            reverse("purchase_differences", kwargs={"period_id": july.id})
        )
        self.assertContains(difference_response, "跨期核销")
        self.assertContains(difference_response, "40.00")

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
        self.assertEqual(sales_invoice.remaining_quantity, Decimal("0"))
        self.assertEqual(purchase_opening.remaining_quantity, Decimal("60"))
        self.assertEqual(goods_opening.remaining_quantity, Decimal("6"))

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
        sales_rows = build_sales_uninvoiced_rows(period)
        self.assertEqual(len(sales_rows), 1)
        self.assertIsNone(sales_rows[0].transaction)
        self.assertEqual(sales_rows[0].outstanding_cost, Decimal("600.00"))
        rollforward = build_goods_issued_rollforward(period)
        self.assertEqual(rollforward["opening_amount"], Decimal("1000.00"))
        self.assertEqual(
            rollforward["invoice_cleared_amount"],
            Decimal("400.00"),
        )
        self.assertEqual(rollforward["ending_amount"], Decimal("600.00"))
        self.assertEqual(rollforward["difference"], Decimal("0.00"))

        self.client.force_login(self.user)
        purchase_page = self.client.get(
            reverse("purchase_matches", kwargs={"period_id": period.id})
        )
        self.assertContains(purchase_page, "期初暂估")
        sales_page = self.client.get(
            reverse("sales_matches", kwargs={"period_id": period.id})
        )
        self.assertContains(sales_page, "期初发出商品")
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

    def test_work_order_costs_page_split_by_workshop(self):
        """工单成本页：每个车间独立一张表（页签切换），不再全车间合并。"""
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
        # 两张独立的表（每车间一张），而不是一张混合表
        self.assertEqual(html.count("<table>"), 2)
        # 表头不再有独立的「车间」列
        self.assertNotIn("<th>车间</th>", html)
        # 「是否用粗纱」列只出现在纺纱车间表头（全表仅 1 处）
        self.assertEqual(html.count("<th>是否用粗纱</th>"), 1)
        # 每车间有合计行
        self.assertIn("纺纱车间合计（1 张工单）", html)
        self.assertIn("机织车间合计（1 张工单）", html)
        # 页签标注工单数
        self.assertIn("纺纱车间 <small>1 张工单</small>", html)
        self.assertIn("机织车间 <small>1 张工单</small>", html)
        # 类别/计划完工/实际完工/完成状态 四列存在
        self.assertIn("<th>类别</th>", html)
        self.assertIn("<th>计划完工时间</th>", html)
        self.assertIn("<th>实际完工时间</th>", html)
        self.assertIn("<th>完成状态</th>", html)
        # 每车间的导出链接带 workshop 参数；顶部"导出Excel（全部车间）"按钮不带 workshop
        export_hrefs = re.findall(
            r'href="([^"]*work_order_cost[^"]*)"', html
        )
        self.assertTrue(export_hrefs)
        self.assertTrue(any("workshop=" in h for h in export_hrefs))
        self.assertIn("导出Excel（全部车间）", html)
        top_href = re.search(r'href="([^"]*work_order_cost[^"]*)"', html).group(1)
        self.assertNotIn("workshop=", top_href)
        # 默认激活第一个车间页签（纺纱）
        self.assertIn('id="ws-tab-0"', html)
        self.assertIn('name="ws-tab"', html)
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
        """历史单价自动补全：期间内新增物料（无上传价）访问单价表时自动按系统计算价补入；
        已上传/已补全的物料不被覆盖。"""
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
        self.assertContains(production_cost_page, "本月生产成本与费用")
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
        self.assertContains(purchase_page, "PI-WB")
        self.assertContains(purchase_page, "GR-WB")
        sales_page = self.client.get(
            reverse("sales_matches", kwargs={"period_id": self.period.id})
        )
        self.assertContains(sales_page, "SI-WB")
        self.assertContains(sales_page, "SO-WB")
        order_page = self.client.get(
            reverse("work_order_matches", kwargs={"period_id": self.period.id})
        )
        self.assertContains(order_page, "MO-WB")
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
        # 三个筛选字段（复选框）+ 口径切换
        for name in ('type="checkbox" name="m"', 'type="checkbox" name="ch"',
                     'type="checkbox" name="cat"'):
            self.assertIn(name, html)
        self.assertIn("按出库期间", html)
        self.assertIn("按开票期间", html)
        # 汇总表在「按月费用录入」之后（原顶端表已删除）
        self.assertLess(html.index("按月费用录入"), html.index("<h2>损益汇总表</h2>"))

    def test_summary_default_shows_current_period_month(self):
        # 打开页面（不带任何筛选）默认只看当月：期间勾当前期间月份，合计按当月重算
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"))
        self._line(2, revenue=Decimal("200"), cost=Decimal("100"), out_month=self.june)
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
                   out_month=date(2026, 5, 1))
        self._line(2, revenue=Decimal("1000"), cost=Decimal("600"), out_month=self.june)
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
        # 「逐行损益」改名为「损益明细表」（行数统计保留）
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"))
        client = self.client
        client.force_login(self.user)
        html = client.get(
            reverse("profit_statement", args=[self.period.id])
        ).content.decode("utf-8")
        self.assertIn("<h2>损益明细表（共 1 行，本页 1 行）</h2>", html)
        self.assertNotIn("逐行损益", html)

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
        "purchase_contracts", "purchase_matches", "purchase_summary",
        "purchase_accruals", "purchase_differences",
        "sales_matches", "sales_uninvoiced", "reconciliation",
        "import_batches", "work_order_costs", "work_order_matches",
        "work_order_report", "work_order_imports", "work_report_imports",
        "material_issue_imports", "bom_imports", "order_completion_report",
        "order_material_report", "price_table", "total_cost",
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
        for name in ("total_cost", "profit_statement", "price_table"):
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

    def test_total_cost_two_row_header_offsets(self):
        # 总成本表：同样是两行表头，第二行贴在第一行（30px）下方
        html = self._html("total_cost")
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
        "purchase_contracts.html", "purchase_differences.html",
        "purchase_matches.html", "purchase_reversal_report.html",
        "sales_matches.html", "sales_uninvoiced.html", "work_order_costs.html",
        "work_order_matches.html", "work_report_imports.html",
        "source_document_detail.html",
    )

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
        self.assertIn("<td><strong>{{ row.material_code }}</strong></td>"
                      "<td>{{ row.material_name }}</td>", tpl)

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
        self.assertIn("<h2>损益明细表（共 2 行，本页 2 行）</h2>", html)
        self.assertIn("台账合计 4 行", html)
        # 筛选控件都在
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
        for name in ("inventory_ledger", "sales_uninvoiced", "work_order_matches",
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
        self.assertEqual(
            heads, ["摘要", "科目", "辅助核算", "借方", "贷方", "数量", "来源行数"]
        )
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
        # 借方按销售大类展开
        categories = {
            row["account"]
            for row in blocks["当月发货当月开票结转成本"]["entries"]
            if row["level"] == "line" and row["debit"] > 0
        }
        self.assertEqual(categories, {"主营业务成本—纱线", "主营业务成本—面料"})
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
        # 借贷两侧逐大类金额一致
        block1 = blocks["当月发货当月开票结转成本"]
        debit_by_cat = {
            row["account"].split("—", 1)[1]: row["debit"]
            for row in block1["entries"]
            if row["level"] == "line" and row["debit"] > 0
        }
        credit_by_cat = {
            row["account"].split("—", 1)[1]: row["credit"]
            for row in block1["entries"]
            if row["level"] == "line" and row["credit"] > 0
        }
        self.assertEqual(debit_by_cat, credit_by_cat)


    def test_voucher_keeps_raw_sales_category(self):
        """销售大类保留台账原值：带括号的写法不能被归一成「其他成品」。"""
        self._line(1, out_month=self.period.month, invoice_month=self.period.month,
                   carry_type="当月发货当月开票结转成本", cost="100.00",
                   category="其他成品（手套、服装、床垫等）")
        self._line(2, out_month=self.period.month, invoice_month=self.period.month,
                   carry_type="当月发货当月开票结转成本", cost="50.00",
                   category="其他成品（毯子）")
        data = build_sales_voucher_entries(self.period)
        accounts = {
            row["account"]
            for block in data["blocks"]
            for row in block["entries"]
            if row["level"] == "line" and row["debit"] > 0
        }
        self.assertEqual(
            accounts,
            {
                "主营业务成本—其他成品（手套、服装、床垫等）",
                "主营业务成本—其他成品（毯子）",
            },
        )
        # 两个原值各自成行（没有被并成一行）
        categories = {
            row["category"]
            for row in data["summary_rows"]
            if row["level"] == "category"
        }
        self.assertEqual(
            categories, {"其他成品（手套、服装、床垫等）", "其他成品（毯子）"}
        )
        # 辅助核算用同一个原值
        aux = {row["aux"] for block in data["blocks"] for row in block["entries"]
               if row["level"] == "line" and row["debit"] > 0}
        self.assertTrue(all("其他成品" in item for item in aux))
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
        self.client.force_login(self.user)
        self.url = reverse("sales_uninvoiced", args=[self.period.id])

    def test_detail_table_splits_issue_month_and_date(self):
        html = self.client.get(self.url).content.decode("utf-8")
        idx = html.index("<h2>截至月末仍未开票的销售出库明细</h2>")
        section = html[html.rindex("<section", 0, idx):]
        section = section[: section.index("</section>")]
        heads = [re.sub(r"<[^>]+>", "", x).strip()
                 for x in re.findall(r"<th[^>]*>(.*?)</th>", section, re.S)]
        self.assertEqual(heads[0], "出库月份")
        self.assertEqual(heads[1], "出库日期")
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

    def test_filter_card_is_removed(self):
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertNotIn('class="section-card filter-card"', html)
        for field in ("issue_month", "document_no", "sales_order_no",
                      "cross_period", 'name="status"'):
            self.assertNotIn(field, html, field)

    def test_other_sections_still_present(self):
        html = self.client.get(self.url).content.decode("utf-8")
        for label in ("发出商品跨月滚动勾稽",
                      "发出商品及销售成本统计表", "截至月末仍未开票的销售出库明细"):
            self.assertIn(label, html)

    def test_voucher_section_moved_to_own_page(self):
        """销售结转成本凭证已拆成独立页：本页只留跳转，独立页有完整凭证表。"""
        html = self.client.get(self.url).content.decode("utf-8")
        self.assertNotIn("<h2>凭证分录（不含税）</h2>", html)
        self.assertNotIn("voucher-filter-form", html)
        self.assertIn(reverse("sales_cost_voucher", args=[self.period.id]), html)

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
        """③ 前期暂估本月到票：借暂估应付 / 贷应付账款，差额进材料成本差异。"""
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
        debit_rows = [
            row for row in block["rows"]
            if row["debit"] > 0 and row["account"].startswith("应付账款")
        ]
        self.assertEqual(
            {row["account"] for row in debit_rows}, {"应付账款—暂估应付—暂估供应商"}
        )
        # 价差单独一行，进材料成本差异
        diff_rows = [row for row in block["rows"] if row["account"] == "材料成本差异"]
        self.assertEqual(len(diff_rows), 1)
        self.assertEqual(diff_rows[0]["debit"], Decimal("50.00"))

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
        self.assertLess(
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
        self.assertIn("原材料—原料", category_accounts)
        self.assertIn("原材料—包装物", category_accounts)
        self.assertEqual(single["totals"]["debit"], category["totals"]["debit"])

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
            "① 本月到货 + 本月到票",
            "② 本月到货 + 月末未到票",
            "③ 前期暂估 + 本月到票",
            "附件1：本月到货明细",
            "附件2：暂估回冲明细",
            "附件3：应付账款—暂估应付余额表",
            "暂估账龄（按到货月份）",
        ):
            self.assertIn(text, html, text)
        self.assertIn("1,000.00", html)

        export = self.client.get(
            reverse("export_listing", args=[self.period.id, "purchase_voucher"])
        )
        self.assertEqual(export.status_code, 200)
        workbook = load_workbook(BytesIO(export.content))
        self.assertEqual(workbook.sheetnames[0], "凭证分录")
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
        self.assertEqual(data["ready_count"], 2)
        self.assertEqual(data["planned_count"], 3)
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
        self.assertEqual(purchase["balance"]["debit"], purchase["balance"]["credit"])
        self.assertTrue(purchase["balance"]["balanced"])
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
            cards["sales"]["entry_headers"],
            ["摘要", "科目", "辅助核算", "借方", "贷方", "数量", "来源行数"],
        )
        self.assertEqual(
            cards["purchase"]["entry_headers"],
            ["摘要", "科目", "辅助核算", "借方", "贷方", "数量", "来源行数"],
        )
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

        self.assertEqual(
            detail_heads,
            ["摘要", "科目", "辅助核算", "借方", "贷方", "数量", "来源行数"],
        )
        self.assertEqual(center_heads, detail_heads)

        # 首行第一列是会计科目（不是结转类型）
        tbody = center_body[center_body.index("<tbody>"):]
        first_tr = tbody[:tbody.index("</tr>")]
        cells = [
            re.sub(r"<[^>]+>", "", c).strip()
            for c in re.findall(r"<td[^>]*>(.*?)</td>", first_tr, re.S)
        ]
        # 中心与明细页同口径（按销售大类展开）：摘要＝结转类型 + 大类，科目带大类
        self.assertTrue(cells[0].startswith("当月发货当月开票结转成本"), cells[0])
        self.assertTrue(cells[1].startswith("主营业务成本—"), cells[1])

    def test_sales_voucher_uses_entry_format_like_purchase(self):
        """销售凭证与采购凭证同格式：汇总卡 + 生成前检查 + 凭证分录（借/贷）+ 附件。"""
        self._receipt()
        self.client.force_login(self.user)
        html = self.client.get(
            reverse("sales_cost_voucher", args=[self.period.id])
        ).content.decode("utf-8")
        for text in ("<h1>销售结转成本凭证</h1>", "summary-strip",
                     "生成前检查", "<h2>凭证分录（不含税）</h2>",
                     "<h2>附件：结转类型 × 销售大类（含表内筛选）</h2>",
                     "取数口径与说明", "voucher-filter-form", "凭证期间"):
            self.assertIn(text, html, text)
        # 采购凭证页也是同一套结构，保证两页格式一致
        purchase = self.client.get(
            reverse("purchase_accrual_voucher", args=[self.period.id])
        ).content.decode("utf-8")
        for text in ("summary-strip", "生成前检查", "<h2>凭证分录（不含税）</h2>"):
            self.assertIn(text, purchase, text)

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
