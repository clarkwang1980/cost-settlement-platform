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
from settlement.services.purchase_accruals import (
    build_purchase_accrual_rows,
    build_purchase_difference_rows,
)
from settlement.services.workorder_costing import (
    build_work_order_cost_summary,
)
from settlement.services.pulp_costing import collect_pulp_costs
from settlement.services.sales_uninvoiced import (
    build_goods_issued_rollforward,
    build_sales_uninvoiced_rows,
)
from settlement.services.profit_statement import (
    ALLOCATION_FIELDS,
    HEADERS,
    build_category_summary,
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

    def test_expense_entry_page_saves_each_month_and_exports(self):
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
            "月度 × 销售渠道", "月份 × 销售渠道 × 销售大类",
        ):
            self.assertContains(response, text)
        # 两个月份各一套输入框
        self.assertContains(response, "exp_2026-06_admin_expense")
        self.assertContains(response, "exp_2026-07_admin_expense")

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
            ["月份", "销售渠道", "销售大类"],
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

        # 起止反了 → 对调成 2026-05 ~ 2026-07
        response = client.get(url + "?basis=out&pfrom=2026-07&pto=2026-05")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.context["period_from"], "2026-05")
        self.assertEqual(response.context["period_to"], "2026-07")

        # 非法值被忽略
        response = client.get(url + "?basis=out&pfrom=abc&pto=2026-05")
        self.assertEqual(response.context["period_from"], "")
        self.assertEqual(response.context["period_to"], "2026-05")

        # 区间只影响顶端表：下方「月×渠道」汇总仍是全部期间
        response = client.get(url + "?basis=out&pfrom=2026-05&pto=2026-05")
        top_months = [r["month"] for r in response.context["category_summary"]["rows"]
                      if r["level"] == "month"]
        all_months = [r["month"] for r in response.context["summary"]["rows"]
                      if r["level"] == "month"]
        self.assertEqual(top_months, ["2026-05"])
        self.assertEqual(all_months, ["2026-05", "2026-07"])
        # 切换口径的链接保留区间参数
        html = response.content.decode("utf-8")
        self.assertIn("pfrom=2026-05", html)

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

    def test_category_filter_on_page_and_isolation(self):
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"), category="纱线")
        self._line(2, revenue=Decimal("500"), cost=Decimal("250"), category="面料")
        client = self.client
        client.force_login(self.user)
        url = reverse("profit_statement", args=[self.period.id])

        # 页面有类别多选与选项（含金额）
        response = client.get(url + "?basis=out")
        self.assertContains(response, 'name="cat"')
        self.assertContains(response, "类别")

        # 多选参数生效
        response = client.get(url + "?basis=out&cat=纱线&cat=面料")
        self.assertEqual(response.context["selected_categories"], ["纱线", "面料"])
        cats = {r["category"] for r in response.context["category_summary"]["rows"]
                if r["level"] == "category"}
        self.assertEqual(cats, {"纱线", "面料"})

        # 不存在的类别被忽略 → 视为不筛选
        response = client.get(url + "?basis=out&cat=不存在")
        self.assertEqual(response.context["selected_categories"], [])
        self.assertFalse(response.context["category_summary"]["has_category_filter"])

        # 只影响顶端表：下方「月×渠道」汇总的合计仍含全部类别
        response = client.get(url + "?basis=out&cat=纱线")
        top = response.context["category_summary"]["grand_total_row"]["revenue"]
        low = response.context["summary"]["grand_total"]["revenue"]
        self.assertEqual(top, Decimal("1000.00"))
        self.assertEqual(low, Decimal("1500.00"))

        # 口径切换链接保留类别参数（URL 编码）
        html = response.content.decode("utf-8")
        self.assertIn("cat=%E7%BA%B1%E7%BA%BF", html)  # “纱线”的 URL 编码
        self.assertIn("显示全部", html)

    def test_expense_entry_page_has_top_summary_table(self):
        self._line(1, revenue=Decimal("1000"), cost=Decimal("600"), category="纱线")
        client = self.client
        client.force_login(self.user)
        url = reverse("profit_statement", args=[self.period.id])
        response = client.get(url)
        self.assertEqual(response.status_code, 200)
        html = response.content.decode("utf-8")
        self.assertContains(response, "汇总：期间 × 销售类别")
        for text in ("销售类别", "销售毛利率(%)", "净利润率(%)"):
            self.assertContains(response, text)
        # 顶端汇总表出现在「按月费用录入」之前
        self.assertLess(html.index("汇总：期间 × 销售类别"), html.index("按月费用录入"))

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
