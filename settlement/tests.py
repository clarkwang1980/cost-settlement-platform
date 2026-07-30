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
    ImportBatch,
    InventoryTransaction,
    MatchAllocation,
    MatchingMode,
    Material,
    OpeningGoodsIssued,
    OpeningInventoryBalance,
    OpeningPurchaseAccrual,
    OpeningWip,
    ProductionOrder,
    PurchaseContractLine,
    PurchaseInvoiceLine,
    ReconciliationResult,
    SalesInvoiceLine,
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
from settlement.services.production_costs import build_production_cost_summary
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
            period, ImportBatch.Kind.PRODUCTION_ORDER, "order-a"
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
        order = ProductionOrder.objects.create(
            period=period,
            batch=order_batch,
            order_no="WO-A",
            workshop="染整",
            product_code=product.code,
            product_name=product.name,
            plan_quantity=Decimal("10"),
            finished_quantity=Decimal("10"),
            status=ProductionOrder.Status.COMPLETED,
            standard_hours=Decimal("1"),
            production_receipt_document="PR-A",
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
            business_type="生产领料",
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
            document_no="PR-A",
            business_type="生产入库",
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
                "车间结转",
                "库存流水",
                "采购匹配",
                "月末采购暂估",
                "采购暂估差异",
                "销售匹配",
                "月末已出库未开票",
                "发出商品滚动",
                "订单成本",
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
            period, ImportBatch.Kind.PRODUCTION_ORDER, "order-chain"
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
        order_a = ProductionOrder.objects.create(
            period=period,
            batch=order_batch,
            order_no="WO-UP",
            workshop="上游车间",
            product_code=semi.code,
            finished_quantity=Decimal("10"),
            plan_quantity=Decimal("10"),
            status=ProductionOrder.Status.COMPLETED,
            production_receipt_document="PR-SEMI",
        )
        order_b = ProductionOrder.objects.create(
            period=period,
            batch=order_batch,
            order_no="WO-DOWN",
            workshop="下游车间",
            product_code=final.code,
            finished_quantity=Decimal("5"),
            plan_quantity=Decimal("5"),
            status=ProductionOrder.Status.COMPLETED,
            production_receipt_document="PR-FINAL",
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
            business_type="生产领料",
            direction=InventoryTransaction.Direction.OUT,
            material=raw,
            quantity=Decimal("40"),
            order_no=order_a.order_no,
        )
        semi_receipt = self.create_txn(
            period,
            inventory_batch,
            txn_date=date(2026, 7, 3),
            document_no="PR-SEMI",
            business_type="生产入库",
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
            business_type="生产领料",
            direction=InventoryTransaction.Direction.OUT,
            material=semi,
            quantity=Decimal("5"),
            order_no=order_b.order_no,
        )
        final_receipt = self.create_txn(
            period,
            inventory_batch,
            txn_date=date(2026, 7, 5),
            document_no="PR-FINAL",
            business_type="生产入库",
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
            june, ImportBatch.Kind.PRODUCTION_ORDER, "roll-order"
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
        order = ProductionOrder.objects.create(
            period=june,
            batch=order_batch,
            order_no="ROLL-WO",
            workshop="跨期车间",
            product_code="ROLL-FG",
            product_name="跨期产品",
            plan_quantity=Decimal("10"),
            finished_quantity=Decimal("0"),
            wip_quantity=Decimal("10"),
            status=ProductionOrder.Status.OPEN,
        )
        self.create_txn(
            june,
            inventory_batch,
            txn_date=date(2026, 6, 15),
            document_no="MI-ROLL",
            business_type="生产领料",
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
        carried_order = ProductionOrder.objects.get(
            period=july, order_no=order.order_no
        )
        self.assertTrue(carried_order.is_system_generated)
        self.assertEqual(carried_order.finished_quantity, Decimal("0"))
        self.assertEqual(carried_order.cost_result.wip_cost, Decimal("40.00"))
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
            period, ImportBatch.Kind.PRODUCTION_ORDER, "order-wip"
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
        order = ProductionOrder.objects.create(
            period=period,
            batch=order_batch,
            order_no="WO-WIP",
            workshop="特纺",
            product_code=product.code,
            product_name=product.name,
            plan_quantity=Decimal("10"),
            finished_quantity=Decimal("8"),
            wip_quantity=Decimal("2"),
            status=ProductionOrder.Status.OPEN,
            standard_hours=Decimal("1"),
            wip_conversion_rate=Decimal("0.5"),
            production_receipt_document="PR-WIP",
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
            business_type="生产领料",
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
            document_no="PR-WIP",
            business_type="生产入库",
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

    def test_ignored_production_order_is_excluded_from_settlement(self):
        period = AccountingPeriod.objects.create(month=date(2026, 10, 1))
        order_batch = self.create_batch(
            period, ImportBatch.Kind.PRODUCTION_ORDER, "ignored-order"
        )
        inventory_batch = self.create_batch(
            period, ImportBatch.Kind.INVENTORY, "ignored-order-inventory"
        )
        product = Material.objects.create(
            code="IGNORED-FG", name="不参与本月核算", uom="件"
        )
        order = ProductionOrder.objects.create(
            period=period,
            batch=order_batch,
            order_no="MO-IGNORED",
            workshop="忽略车间",
            product_code="IGNORED-FG",
            product_name="不参与本月核算",
            plan_quantity=Decimal("10"),
            finished_quantity=Decimal("10"),
            wip_quantity=Decimal("0"),
            status=ProductionOrder.Status.COMPLETED,
            match_mode=MatchingMode.IGNORED,
            ignore_reason="非成本核算订单",
        )
        receipt = self.create_txn(
            period,
            inventory_batch,
            txn_date=date(2026, 10, 20),
            document_no="PR-IGNORED",
            business_type="生产入库",
            direction=InventoryTransaction.Direction.IN,
            material=product,
            quantity=Decimal("10"),
            order_no=order.order_no,
            workshop=order.workshop,
        )

        run = run_settlement(period)

        self.assertTrue(run.converged)
        receipt.refresh_from_db()
        self.assertEqual(
            receipt.match_status,
            InventoryTransaction.MatchStatus.NOT_REQUIRED,
        )
        self.assertFalse(receipt.allocations.exists())
        self.assertFalse(period.order_cost_results.filter(order=order).exists())
        self.assertFalse(
            period.workshop_settlements.filter(workshop="忽略车间").exists()
        )
        self.assertFalse(
            period.exceptions.filter(reference=order.order_no).exists()
        )
        self.assertFalse(
            period.exceptions.filter(reference__contains="PR-IGNORED").exists()
        )


class ImportAndWebTests(SettlementFixtureMixin, TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="web-user", password="test-password"
        )
        self.period = AccountingPeriod.objects.create(month=date(2026, 8, 1))

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
        self.assertContains(upload_page, "九类资料字段说明")
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

    def test_all_supplied_templates_import_and_close_end_to_end(self):
        template_map = [
            (ImportBatch.Kind.INVENTORY, "库存出入库流水.csv"),
            (ImportBatch.Kind.PURCHASE_CONTRACT, "采购合同.csv"),
            (ImportBatch.Kind.PURCHASE_INVOICE, "采购发票.csv"),
            (ImportBatch.Kind.SALES_INVOICE, "销售发票.csv"),
            (ImportBatch.Kind.PRODUCTION_ORDER, "生产订单.csv"),
            (ImportBatch.Kind.WORKSHOP_COST, "车间费用.csv"),
            (ImportBatch.Kind.OPENING_WIP, "期初在产品.csv"),
            (
                ImportBatch.Kind.OPENING_PURCHASE_ACCRUAL,
                "期初采购暂估明细.csv",
            ),
            (
                ImportBatch.Kind.OPENING_GOODS_ISSUED,
                "期初发出商品明细.csv",
            ),
        ]
        template_dir = Path(settings.BASE_DIR) / "sample_templates"
        with TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                for kind, filename in template_map:
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
        self.assertEqual(self.period.import_batches.count(), 9)
        self.assertEqual(self.period.inventory_transactions.count(), 4)
        self.assertEqual(self.period.production_orders.count(), 1)
        self.assertEqual(self.period.order_cost_results.count(), 1)
        result = self.period.order_cost_results.get()
        self.assertEqual(result.finished_cost, Decimal("650.00"))
        self.assertEqual(result.wip_cost, Decimal("0.00"))
        self.assertFalse(
            self.period.reconciliations.filter(
                status=ReconciliationResult.Status.FAIL
            ).exists()
        )
        summary = build_production_cost_summary(self.period)
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
        self.assertContains(workshop_page, "染整订单成本明细")
        self.assertContains(workshop_page, "WO-001")

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
            ImportBatch.Kind.PRODUCTION_ORDER,
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
        order = ProductionOrder.objects.create(
            period=self.period,
            batch=order_batch,
            order_no="MO-WB",
            workshop="测试车间",
            product_code=product.code,
            product_name=product.name,
            plan_quantity=Decimal("5"),
            finished_quantity=Decimal("5"),
            wip_quantity=Decimal("0"),
            status=ProductionOrder.Status.COMPLETED,
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
            reverse("order_matches", kwargs={"period_id": self.period.id})
        )
        self.assertContains(order_page, "MO-WB")
        self.assertContains(order_page, "ISS-WB")
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
        self.assertEqual(order.match_mode, MatchingMode.AUTO)
        self.assertEqual(material_issue.match_mode, MatchingMode.AUTO)
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
