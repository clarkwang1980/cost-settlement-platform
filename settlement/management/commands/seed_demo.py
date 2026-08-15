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
    help = "建立一套可以直接查看和测试的演示月结数据（默认 2026-07，可用 --month 指定）。"

    def add_arguments(self, parser):
        parser.add_argument(
            "--month",
            default="2026-07",
            help="核算月份，格式 YYYY-MM，例如 2026-07",
        )

    def handle(self, *args, **options):
        year, month_no = (int(part) for part in options["month"].split("-"))
        month = date(year, month_no, 1)
        month_code = month.strftime("%Y%m")

        user_model = get_user_model()
        user, created = user_model.objects.get_or_create(
            username="admin",
            defaults={"is_staff": True, "is_superuser": True},
        )
        if created:
            user.set_password("admin123")
            user.save()

        period, _ = AccountingPeriod.objects.get_or_create(month=month)
        if period.inventory_transactions.exists():
            self.stdout.write(f"{period} 演示数据已经存在。")
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

        raw, _ = Material.objects.get_or_create(
            code="RM-001", defaults={"name": "演示原料", "uom": "kg"}
        )
        product, _ = Material.objects.get_or_create(
            code="FG-001", defaults={"name": "演示成品", "uom": "kg"}
        )

        PurchaseContractLine.objects.create(
            period=period,
            batch=contract_batch,
            contract_no=f"PC-{month_code}-001",
            contract_date=month.replace(day=2),
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
            invoice_no=f"PI-{month_code}-001",
            invoice_date=month.replace(day=25),
            invoice_period=month,
            contract_no=f"PC-{month_code}-001",
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
            invoice_no=f"SI-{month_code}-001",
            invoice_date=month.replace(day=28),
            invoice_period=month,
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
            order_no=f"WO-{month_code}-001",
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
            txn_date=month.replace(day=5),
            document_no="GR-001",
            business_type="采购入库",
            direction=InventoryTransaction.Direction.IN,
            material=raw,
            material_code=raw.code,
            material_name=raw.name,
            quantity=Decimal("100"),
            supplier_code="SUP-001",
            supplier_name="演示供应商",
            contract_no=f"PC-{month_code}-001",
        )
        inventory(
            txn_date=month.replace(day=10),
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
            txn_date=month.replace(day=20),
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
            txn_date=month.replace(day=24),
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
                f"{period} 演示月结完成：{run.iterations}次循环，"
                f"{'已收敛' if run.converged else '未收敛'}。"
            )
        )
        self.stdout.write("演示账号：admin / admin123")
