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
    ReconciliationResult,
    SalesInvoiceLine,
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
from settlement.services.sales_uninvoiced import (
    build_goods_issued_rollforward,
    build_sales_uninvoiced_rows,
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
        # 每个车间的导出链接带 workshop 参数
        export_hrefs = re.findall(
            r'href="([^"]*work_order_cost[^"]*)"', html
        )
        self.assertTrue(export_hrefs)
        self.assertTrue(all("workshop=" in h for h in export_hrefs))
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
        # 导出含上传价
        resp = self.client.get(
            reverse("price_table", args=[self.period.id]) + "?export=1"
        )
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(BytesIO(resp.content))
        ws = wb.active
        found = False
        for row in ws.iter_rows(values_only=True):
            if row and row[1] == "11.01":
                found = True
                self.assertEqual(float(row[4]), 88.88)
                self.assertEqual(float(row[5]), 77.77)
                break
        self.assertTrue(found)

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
        self.assertContains(upload_page, "十七类资料字段说明")
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
            (ImportBatch.Kind.OPENING_WIP, "期初在产品.csv"),
            (ImportBatch.Kind.OPENING_INVENTORY, "期初存货.csv"),
            (ImportBatch.Kind.ENDING_INVENTORY, "期末存货.csv"),
            (
                ImportBatch.Kind.OPENING_PURCHASE_ACCRUAL,
                "期初采购暂估明细.csv",
            ),
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
        # 系统不再自动生成期末行
        self.assertFalse(
            self.period.ending_inventory_balances.filter(
                is_system_generated=True
            ).exists()
        )
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
        # 期末存货以用户上传报表为准：未上传期末时，系统不自动生成期末行
        self.assertFalse(
            EndingInventoryBalance.objects.filter(
                period=period, is_system_generated=True
            ).exists()
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
