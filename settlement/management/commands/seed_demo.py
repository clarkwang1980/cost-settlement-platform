from datetime import date
from decimal import Decimal

from django.contrib.auth import get_user_model
from django.core.files.base import ContentFile
from django.core.management.base import BaseCommand

from settlement.models import (
    AccountingPeriod,
    ImportBatch,
    InventoryTransaction,
    Material,
    ProductionOrder,
    PurchaseContractLine,
    PurchaseInvoiceLine,
    SalesInvoiceLine,
    WorkshopCostPool,
)
from settlement.services.costing import run_settlement


class Command(BaseCommand):
    help = "建立一套可以直接查看和测试的演示月结数据。"

    def handle(self, *args, **options):
        user_model = get_user_model()
        user, created = user_model.objects.get_or_create(
            username="admin",
            defaults={"is_staff": True, "is_superuser": True},
        )
        if created:
            user.set_password("admin123")
            user.save()

        period, _ = AccountingPeriod.objects.get_or_create(month=date(2026, 6, 1))
        if period.inventory_transactions.exists():
            self.stdout.write("演示数据已经存在。")
            return

        def batch(kind):
            item = ImportBatch.objects.create(
                period=period,
                kind=kind,
                original_name=f"demo-{kind}.csv",
                sha256=f"{kind:0<64}"[:64],
                status=ImportBatch.Status.SUCCESS,
                row_count=1,
                uploaded_by=user,
            )
            item.file.save(
                f"demo-{kind}.csv",
                ContentFile("演示数据".encode("utf-8")),
                save=True,
            )
            return item

        inventory_batch = batch(ImportBatch.Kind.INVENTORY)
        contract_batch = batch(ImportBatch.Kind.PURCHASE_CONTRACT)
        purchase_invoice_batch = batch(ImportBatch.Kind.PURCHASE_INVOICE)
        sales_invoice_batch = batch(ImportBatch.Kind.SALES_INVOICE)
        order_batch = batch(ImportBatch.Kind.PRODUCTION_ORDER)
        cost_batch = batch(ImportBatch.Kind.WORKSHOP_COST)

        raw = Material.objects.create(code="RM-001", name="演示原料", uom="kg")
        product = Material.objects.create(code="FG-001", name="演示成品", uom="kg")

        PurchaseContractLine.objects.create(
            period=period,
            batch=contract_batch,
            contract_no="PC-202606-001",
            contract_date=date(2026, 6, 2),
            supplier_code="SUP-001",
            supplier_name="演示供应商",
            material_code=raw.code,
            material_name=raw.name,
            quantity=Decimal("100"),
            remaining_quantity=Decimal("100"),
            unit_price_excl_tax=Decimal("10"),
        )
        PurchaseInvoiceLine.objects.create(
            period=period,
            batch=purchase_invoice_batch,
            invoice_no="PI-202606-001",
            invoice_date=date(2026, 6, 25),
            invoice_period=date(2026, 6, 1),
            contract_no="PC-202606-001",
            supplier_code="SUP-001",
            supplier_name="演示供应商",
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
            invoice_no="SI-202606-001",
            invoice_date=date(2026, 6, 28),
            invoice_period=date(2026, 6, 1),
            sales_order_no="SO-001",
            customer_code="CUS-001",
            customer_name="演示客户",
            material_code=product.code,
            material_name=product.name,
            quantity=Decimal("4"),
            remaining_quantity=Decimal("4"),
        )
        order = ProductionOrder.objects.create(
            period=period,
            batch=order_batch,
            order_no="WO-202606-001",
            workshop="染整",
            product_code=product.code,
            product_name=product.name,
            plan_quantity=Decimal("10"),
            finished_quantity=Decimal("10"),
            wip_quantity=Decimal("0"),
            status=ProductionOrder.Status.COMPLETED,
            standard_hours=Decimal("1"),
            wip_conversion_rate=Decimal("1"),
            production_receipt_document="PR-001",
        )
        WorkshopCostPool.objects.create(
            period=period,
            batch=cost_batch,
            workshop="染整",
            water_power=Decimal("100"),
            wages=Decimal("50"),
        )

        def inventory(**kwargs):
            return InventoryTransaction.objects.create(
                period=period,
                batch=inventory_batch,
                line_no=str(
                    InventoryTransaction.objects.filter(period=period).count() + 1
                ),
                warehouse="演示仓库",
                specification="",
                uom="kg",
                match_status=InventoryTransaction.MatchStatus.UNMATCHED,
                **kwargs,
            )

        inventory(
            txn_date=date(2026, 6, 5),
            document_no="GR-001",
            business_type="采购入库",
            direction=InventoryTransaction.Direction.IN,
            material=raw,
            material_code=raw.code,
            material_name=raw.name,
            quantity=Decimal("100"),
            supplier_code="SUP-001",
            supplier_name="演示供应商",
            contract_no="PC-202606-001",
        )
        inventory(
            txn_date=date(2026, 6, 10),
            document_no="MI-001",
            business_type="生产领料",
            direction=InventoryTransaction.Direction.OUT,
            material=raw,
            material_code=raw.code,
            material_name=raw.name,
            quantity=Decimal("50"),
            order_no=order.order_no,
            workshop="染整",
        )
        inventory(
            txn_date=date(2026, 6, 20),
            document_no="PR-001",
            business_type="生产入库",
            direction=InventoryTransaction.Direction.IN,
            material=product,
            material_code=product.code,
            material_name=product.name,
            quantity=Decimal("10"),
            order_no=order.order_no,
            workshop="染整",
        )
        inventory(
            txn_date=date(2026, 6, 24),
            document_no="SO-001",
            business_type="销售出库",
            direction=InventoryTransaction.Direction.OUT,
            material=product,
            material_code=product.code,
            material_name=product.name,
            quantity=Decimal("4"),
            order_no="SO-001",
            customer_code="CUS-001",
            customer_name="演示客户",
        )

        run = run_settlement(period)
        self.stdout.write(
            self.style.SUCCESS(
                f"演示月结完成：{run.iterations}次循环，"
                f"{'已收敛' if run.converged else '未收敛'}。"
            )
        )
        self.stdout.write("演示账号：admin / admin123")

