from decimal import Decimal
from django.conf import settings
from django.db import models


ZERO = Decimal("0")


class MatchingMode(models.TextChoices):
    AUTO = "auto", "自动匹配"
    MANUAL = "manual", "手工匹配"
    IGNORED = "ignored", "忽略匹配"


class AccountingPeriod(models.Model):
    class Status(models.TextChoices):
        OPEN = "open", "待处理"
        CALCULATED = "calculated", "已计算"
        REVIEWED = "reviewed", "已复核"
        LOCKED = "locked", "已锁账"

    month = models.DateField(unique=True, verbose_name="核算月份")
    status = models.CharField(
        max_length=20, choices=Status.choices, default=Status.OPEN, verbose_name="状态"
    )
    last_run_at = models.DateTimeField(null=True, blank=True, verbose_name="最后结算时间")
    locked_at = models.DateTimeField(null=True, blank=True, verbose_name="锁账时间")
    locked_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        null=True,
        blank=True,
        on_delete=models.SET_NULL,
        related_name="locked_periods",
        verbose_name="锁账人",
    )
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-month"]
        verbose_name = "核算期间"
        verbose_name_plural = "核算期间"

    def __str__(self):
        return self.month.strftime("%Y年%m月")


class ImportBatch(models.Model):
    class Kind(models.TextChoices):
        INVENTORY = "inventory", "库存出入库流水"
        PURCHASE_CONTRACT = "purchase_contract", "采购合同"
        PURCHASE_INVOICE = "purchase_invoice", "采购发票"
        SALES_INVOICE = "sales_invoice", "销售发票"
        PRODUCTION_ORDER = "production_order", "生产订单"
        WORKSHOP_COST = "workshop_cost", "车间费用"
        OPENING_WIP = "opening_wip", "期初在产品"
        OPENING_PURCHASE_ACCRUAL = (
            "opening_purchase_accrual",
            "期初采购暂估明细",
        )
        OPENING_GOODS_ISSUED = (
            "opening_goods_issued",
            "期初发出商品明细",
        )

    class Status(models.TextChoices):
        PENDING = "pending", "等待处理"
        SUCCESS = "success", "导入成功"
        FAILED = "failed", "导入失败"

    period = models.ForeignKey(
        AccountingPeriod, on_delete=models.CASCADE, related_name="import_batches"
    )
    kind = models.CharField(max_length=30, choices=Kind.choices)
    file = models.FileField(upload_to="imports/%Y/%m/")
    original_name = models.CharField(max_length=255)
    sha256 = models.CharField(max_length=64, db_index=True)
    status = models.CharField(
        max_length=20, choices=Status.choices, default=Status.PENDING
    )
    row_count = models.PositiveIntegerField(default=0)
    error_count = models.PositiveIntegerField(default=0)
    message = models.TextField(blank=True)
    uploaded_by = models.ForeignKey(
        settings.AUTH_USER_MODEL, null=True, on_delete=models.SET_NULL
    )
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-created_at"]
        constraints = [
            models.UniqueConstraint(
                fields=["period", "kind", "sha256"], name="unique_period_kind_file_hash"
            )
        ]

    def __str__(self):
        return f"{self.period} - {self.get_kind_display()} - {self.original_name}"


class Material(models.Model):
    code = models.CharField(max_length=80, unique=True, verbose_name="物料编码")
    name = models.CharField(max_length=255, blank=True, verbose_name="物料名称")
    specification = models.CharField(max_length=255, blank=True, verbose_name="规格")
    uom = models.CharField(max_length=30, blank=True, verbose_name="单位")

    class Meta:
        ordering = ["code"]
        verbose_name = "物料"
        verbose_name_plural = "物料"

    def __str__(self):
        return f"{self.code} {self.name}".strip()


class InventoryTransaction(models.Model):
    class Direction(models.TextChoices):
        IN = "in", "入库"
        OUT = "out", "出库"

    class MatchStatus(models.TextChoices):
        UNMATCHED = "unmatched", "未匹配"
        PARTIAL = "partial", "部分匹配"
        MATCHED = "matched", "已匹配"
        NOT_REQUIRED = "not_required", "无需匹配"

    period = models.ForeignKey(
        AccountingPeriod, on_delete=models.CASCADE, related_name="inventory_transactions"
    )
    batch = models.ForeignKey(
        ImportBatch, on_delete=models.CASCADE, related_name="inventory_transactions"
    )
    txn_date = models.DateField(verbose_name="业务日期")
    document_no = models.CharField(max_length=100, blank=True, db_index=True)
    line_no = models.CharField(max_length=40, blank=True)
    business_type = models.CharField(max_length=80, verbose_name="业务类型")
    direction = models.CharField(max_length=5, choices=Direction.choices)
    warehouse = models.CharField(max_length=100, blank=True)
    material = models.ForeignKey(
        Material, null=True, on_delete=models.SET_NULL, related_name="transactions"
    )
    material_code = models.CharField(max_length=80, db_index=True)
    material_name = models.CharField(max_length=255, blank=True)
    specification = models.CharField(max_length=255, blank=True)
    quantity = models.DecimalField(max_digits=20, decimal_places=6)
    uom = models.CharField(max_length=30, blank=True)
    order_no = models.CharField(max_length=100, blank=True, db_index=True)
    workshop = models.CharField(max_length=100, blank=True)
    supplier_code = models.CharField(max_length=80, blank=True)
    supplier_name = models.CharField(max_length=255, blank=True)
    customer_code = models.CharField(max_length=80, blank=True)
    customer_name = models.CharField(max_length=255, blank=True)
    contract_no = models.CharField(max_length=100, blank=True)
    invoice_no = models.CharField(max_length=100, blank=True)
    unit_price = models.DecimalField(
        max_digits=20, decimal_places=8, null=True, blank=True
    )
    amount = models.DecimalField(
        max_digits=20, decimal_places=2, null=True, blank=True
    )
    price_source = models.CharField(max_length=40, blank=True)
    invoice_period = models.DateField(null=True, blank=True)
    match_status = models.CharField(
        max_length=20,
        choices=MatchStatus.choices,
        default=MatchStatus.UNMATCHED,
    )
    match_mode = models.CharField(
        max_length=20,
        choices=MatchingMode.choices,
        default=MatchingMode.AUTO,
    )
    ignore_reason = models.CharField(max_length=500, blank=True)
    remark = models.CharField(max_length=500, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["txn_date", "document_no", "line_no", "id"]
        indexes = [
            models.Index(fields=["period", "material_code", "direction"]),
            models.Index(fields=["period", "business_type"]),
            models.Index(fields=["period", "order_no"]),
        ]

    def __str__(self):
        return f"{self.txn_date} {self.document_no} {self.material_code}"


class PurchaseContractLine(models.Model):
    period = models.ForeignKey(
        AccountingPeriod, on_delete=models.CASCADE, related_name="purchase_contracts"
    )
    batch = models.ForeignKey(
        ImportBatch,
        null=True,
        blank=True,
        on_delete=models.CASCADE,
        related_name="purchase_contracts",
    )
    carried_from = models.ForeignKey(
        "self",
        null=True,
        blank=True,
        on_delete=models.PROTECT,
        related_name="carried_contract_lines",
    )
    is_system_generated = models.BooleanField(default=False)
    contract_no = models.CharField(max_length=100, db_index=True)
    contract_date = models.DateField(null=True, blank=True)
    supplier_code = models.CharField(max_length=80, blank=True)
    supplier_name = models.CharField(max_length=255, blank=True)
    material_code = models.CharField(max_length=80, db_index=True)
    material_name = models.CharField(max_length=255, blank=True)
    quantity = models.DecimalField(max_digits=20, decimal_places=6)
    unit_price_excl_tax = models.DecimalField(max_digits=20, decimal_places=8)
    remaining_quantity = models.DecimalField(max_digits=20, decimal_places=6)

    class Meta:
        ordering = ["contract_date", "contract_no", "id"]


class PurchaseInvoiceLine(models.Model):
    period = models.ForeignKey(
        AccountingPeriod, on_delete=models.CASCADE, related_name="purchase_invoices"
    )
    batch = models.ForeignKey(
        ImportBatch, on_delete=models.CASCADE, related_name="purchase_invoices"
    )
    invoice_no = models.CharField(max_length=100, db_index=True)
    invoice_date = models.DateField()
    invoice_period = models.DateField()
    contract_no = models.CharField(max_length=100, blank=True)
    supplier_code = models.CharField(max_length=80, blank=True)
    supplier_name = models.CharField(max_length=255, blank=True)
    material_code = models.CharField(max_length=80, db_index=True)
    material_name = models.CharField(max_length=255, blank=True)
    quantity = models.DecimalField(max_digits=20, decimal_places=6)
    unit_price_excl_tax = models.DecimalField(max_digits=20, decimal_places=8)
    amount_excl_tax = models.DecimalField(max_digits=20, decimal_places=2)
    remaining_quantity = models.DecimalField(max_digits=20, decimal_places=6)
    match_mode = models.CharField(
        max_length=20,
        choices=MatchingMode.choices,
        default=MatchingMode.AUTO,
    )
    ignore_reason = models.CharField(max_length=500, blank=True)

    class Meta:
        ordering = ["invoice_date", "invoice_no", "id"]


class SalesInvoiceLine(models.Model):
    period = models.ForeignKey(
        AccountingPeriod, on_delete=models.CASCADE, related_name="sales_invoices"
    )
    batch = models.ForeignKey(
        ImportBatch, on_delete=models.CASCADE, related_name="sales_invoices"
    )
    invoice_no = models.CharField(max_length=100, db_index=True)
    invoice_date = models.DateField()
    invoice_period = models.DateField()
    sales_order_no = models.CharField(max_length=100, blank=True)
    customer_code = models.CharField(max_length=80, blank=True)
    customer_name = models.CharField(max_length=255, blank=True)
    material_code = models.CharField(max_length=80, db_index=True)
    material_name = models.CharField(max_length=255, blank=True)
    quantity = models.DecimalField(max_digits=20, decimal_places=6)
    unit_price_excl_tax = models.DecimalField(
        max_digits=20, decimal_places=8, null=True, blank=True
    )
    amount_excl_tax = models.DecimalField(
        max_digits=20, decimal_places=2, null=True, blank=True
    )
    remaining_quantity = models.DecimalField(max_digits=20, decimal_places=6)
    match_mode = models.CharField(
        max_length=20,
        choices=MatchingMode.choices,
        default=MatchingMode.AUTO,
    )
    ignore_reason = models.CharField(max_length=500, blank=True)

    class Meta:
        ordering = ["invoice_date", "invoice_no", "id"]


class ProductionOrder(models.Model):
    class Status(models.TextChoices):
        OPEN = "open", "未完工"
        COMPLETED = "completed", "完工"

    period = models.ForeignKey(
        AccountingPeriod, on_delete=models.CASCADE, related_name="production_orders"
    )
    batch = models.ForeignKey(
        ImportBatch,
        null=True,
        blank=True,
        on_delete=models.CASCADE,
        related_name="production_orders",
    )
    carried_from = models.ForeignKey(
        "self",
        null=True,
        blank=True,
        on_delete=models.PROTECT,
        related_name="carried_orders",
    )
    is_system_generated = models.BooleanField(default=False)
    order_no = models.CharField(max_length=100, db_index=True)
    workshop = models.CharField(max_length=100, db_index=True)
    product_code = models.CharField(max_length=80, db_index=True)
    product_name = models.CharField(max_length=255, blank=True)
    plan_quantity = models.DecimalField(max_digits=20, decimal_places=6, default=ZERO)
    finished_quantity = models.DecimalField(max_digits=20, decimal_places=6, default=ZERO)
    wip_quantity = models.DecimalField(max_digits=20, decimal_places=6, default=ZERO)
    status = models.CharField(max_length=20, choices=Status.choices, default=Status.OPEN)
    standard_hours = models.DecimalField(
        max_digits=20, decimal_places=6, default=Decimal("1")
    )
    material_completion_rate = models.DecimalField(
        max_digits=8, decimal_places=6, null=True, blank=True
    )
    wip_conversion_rate = models.DecimalField(
        max_digits=8, decimal_places=6, default=Decimal("0.5")
    )
    production_receipt_document = models.CharField(max_length=100, blank=True)
    match_mode = models.CharField(
        max_length=20,
        choices=MatchingMode.choices,
        default=MatchingMode.AUTO,
    )
    ignore_reason = models.CharField(max_length=500, blank=True)

    class Meta:
        ordering = ["workshop", "order_no"]
        constraints = [
            models.UniqueConstraint(
                fields=["period", "order_no"], name="unique_order_per_period"
            )
        ]

    def __str__(self):
        return f"{self.order_no} {self.product_code}"


class WorkshopCostPool(models.Model):
    period = models.ForeignKey(
        AccountingPeriod, on_delete=models.CASCADE, related_name="workshop_cost_pools"
    )
    batch = models.ForeignKey(
        ImportBatch, on_delete=models.CASCADE, related_name="workshop_cost_pools"
    )
    workshop = models.CharField(max_length=100)
    water_power = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    steam = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    natural_gas = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    depreciation = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    consumables = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    wages = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=["period", "workshop"], name="unique_cost_pool_per_workshop"
            )
        ]


class WorkshopSettlement(models.Model):
    class Status(models.TextChoices):
        PENDING = "pending", "待结转"
        SETTLED = "settled", "已结转"
        FAILED = "failed", "结转失败"

    period = models.ForeignKey(
        AccountingPeriod,
        on_delete=models.CASCADE,
        related_name="workshop_settlements",
    )
    workshop = models.CharField(max_length=100, db_index=True)
    sequence = models.PositiveIntegerField(default=1)
    status = models.CharField(
        max_length=20, choices=Status.choices, default=Status.PENDING
    )
    iterations = models.PositiveIntegerField(default=0)
    converged = models.BooleanField(default=False)
    settled_at = models.DateTimeField(null=True, blank=True)
    settled_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        null=True,
        blank=True,
        on_delete=models.SET_NULL,
        related_name="settled_workshops",
    )
    order_cost = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    production_receipt_amount = models.DecimalField(
        max_digits=20, decimal_places=2, default=ZERO
    )
    error_message = models.TextField(blank=True)

    class Meta:
        ordering = ["sequence", "workshop"]
        constraints = [
            models.UniqueConstraint(
                fields=["period", "workshop"],
                name="unique_workshop_settlement_per_period",
            )
        ]

    def __str__(self):
        return f"{self.period} {self.sequence} {self.workshop}"


class OpeningWip(models.Model):
    period = models.ForeignKey(
        AccountingPeriod, on_delete=models.CASCADE, related_name="opening_wip"
    )
    batch = models.ForeignKey(
        ImportBatch,
        null=True,
        blank=True,
        on_delete=models.CASCADE,
        related_name="opening_wip",
    )
    source_period = models.ForeignKey(
        AccountingPeriod,
        null=True,
        blank=True,
        on_delete=models.PROTECT,
        related_name="carried_wip_balances",
    )
    is_system_generated = models.BooleanField(default=False)
    order_no = models.CharField(max_length=100, db_index=True)
    material = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    water_power = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    steam = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    natural_gas = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    depreciation = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    consumables = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    wages = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=["period", "order_no"], name="unique_opening_wip_per_order"
            )
        ]

    @property
    def total(self):
        return (
            self.material
            + self.water_power
            + self.steam
            + self.natural_gas
            + self.depreciation
            + self.consumables
            + self.wages
        )


class OpeningInventoryBalance(models.Model):
    period = models.ForeignKey(
        AccountingPeriod,
        on_delete=models.CASCADE,
        related_name="opening_inventory_balances",
    )
    source_period = models.ForeignKey(
        AccountingPeriod,
        on_delete=models.PROTECT,
        related_name="carried_inventory_balances",
    )
    material_code = models.CharField(max_length=80, db_index=True)
    material_name = models.CharField(max_length=255, blank=True)
    quantity = models.DecimalField(max_digits=20, decimal_places=6, default=ZERO)
    amount = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    unit_price = models.DecimalField(max_digits=20, decimal_places=8, default=ZERO)
    is_system_generated = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["material_code"]
        constraints = [
            models.UniqueConstraint(
                fields=["period", "material_code"],
                name="unique_opening_inventory_per_material",
            )
        ]

    def __str__(self):
        return f"{self.period} {self.material_code} {self.quantity}"


class OpeningPurchaseAccrual(models.Model):
    period = models.ForeignKey(
        AccountingPeriod,
        on_delete=models.CASCADE,
        related_name="opening_purchase_accruals",
    )
    batch = models.ForeignKey(
        ImportBatch,
        on_delete=models.CASCADE,
        related_name="opening_purchase_accruals",
    )
    original_period = models.DateField()
    document_date = models.DateField()
    document_no = models.CharField(max_length=100, db_index=True)
    line_no = models.CharField(max_length=40, blank=True)
    supplier_code = models.CharField(max_length=80, blank=True)
    supplier_name = models.CharField(max_length=255, blank=True)
    contract_no = models.CharField(max_length=100, blank=True)
    material_code = models.CharField(max_length=80, db_index=True)
    material_name = models.CharField(max_length=255, blank=True)
    uom = models.CharField(max_length=30, blank=True)
    original_quantity = models.DecimalField(max_digits=20, decimal_places=6)
    prior_invoiced_quantity = models.DecimalField(
        max_digits=20,
        decimal_places=6,
        default=ZERO,
    )
    opening_quantity = models.DecimalField(max_digits=20, decimal_places=6)
    provisional_unit_price = models.DecimalField(
        max_digits=20,
        decimal_places=8,
    )
    opening_amount = models.DecimalField(max_digits=20, decimal_places=2)
    remaining_quantity = models.DecimalField(max_digits=20, decimal_places=6)
    invoice_period = models.DateField(null=True, blank=True)
    match_status = models.CharField(
        max_length=20,
        choices=InventoryTransaction.MatchStatus.choices,
        default=InventoryTransaction.MatchStatus.UNMATCHED,
    )
    match_mode = models.CharField(
        max_length=20,
        choices=MatchingMode.choices,
        default=MatchingMode.AUTO,
    )
    ignore_reason = models.CharField(max_length=500, blank=True)
    remark = models.CharField(max_length=500, blank=True)

    class Meta:
        ordering = [
            "original_period",
            "document_date",
            "document_no",
            "line_no",
            "id",
        ]
        constraints = [
            models.UniqueConstraint(
                fields=[
                    "period",
                    "document_no",
                    "line_no",
                    "material_code",
                ],
                name="unique_opening_purchase_accrual_line",
            )
        ]

    def __str__(self):
        return f"{self.document_no} {self.material_code} {self.opening_quantity}"


class OpeningGoodsIssued(models.Model):
    period = models.ForeignKey(
        AccountingPeriod,
        on_delete=models.CASCADE,
        related_name="opening_goods_issued",
    )
    batch = models.ForeignKey(
        ImportBatch,
        on_delete=models.CASCADE,
        related_name="opening_goods_issued",
    )
    original_period = models.DateField()
    document_date = models.DateField()
    document_no = models.CharField(max_length=100, db_index=True)
    line_no = models.CharField(max_length=40, blank=True)
    sales_order_no = models.CharField(max_length=100, blank=True, db_index=True)
    customer_code = models.CharField(max_length=80, blank=True)
    customer_name = models.CharField(max_length=255, blank=True)
    material_code = models.CharField(max_length=80, db_index=True)
    material_name = models.CharField(max_length=255, blank=True)
    uom = models.CharField(max_length=30, blank=True)
    original_quantity = models.DecimalField(max_digits=20, decimal_places=6)
    prior_invoiced_quantity = models.DecimalField(
        max_digits=20,
        decimal_places=6,
        default=ZERO,
    )
    opening_quantity = models.DecimalField(max_digits=20, decimal_places=6)
    unit_cost = models.DecimalField(max_digits=20, decimal_places=8)
    opening_amount = models.DecimalField(max_digits=20, decimal_places=2)
    remaining_quantity = models.DecimalField(max_digits=20, decimal_places=6)
    invoice_period = models.DateField(null=True, blank=True)
    match_status = models.CharField(
        max_length=20,
        choices=InventoryTransaction.MatchStatus.choices,
        default=InventoryTransaction.MatchStatus.UNMATCHED,
    )
    match_mode = models.CharField(
        max_length=20,
        choices=MatchingMode.choices,
        default=MatchingMode.AUTO,
    )
    ignore_reason = models.CharField(max_length=500, blank=True)
    remark = models.CharField(max_length=500, blank=True)

    class Meta:
        ordering = [
            "original_period",
            "document_date",
            "document_no",
            "line_no",
            "id",
        ]
        constraints = [
            models.UniqueConstraint(
                fields=[
                    "period",
                    "document_no",
                    "line_no",
                    "material_code",
                ],
                name="unique_opening_goods_issued_line",
            )
        ]

    def __str__(self):
        return f"{self.document_no} {self.material_code} {self.opening_quantity}"


class MatchAllocation(models.Model):
    class Kind(models.TextChoices):
        PURCHASE_CONTRACT = "purchase_contract", "采购入库-合同"
        PURCHASE_INVOICE = "purchase_invoice", "采购入库-发票"
        SALES_INVOICE = "sales_invoice", "销售出库-发票"
        ORDER_MATERIAL = "order_material", "生产领退料-订单"
        PRODUCTION_RECEIPT = "production_receipt", "生产入库-订单"

    period = models.ForeignKey(
        AccountingPeriod, on_delete=models.CASCADE, related_name="match_allocations"
    )
    kind = models.CharField(max_length=30, choices=Kind.choices)
    inventory_transaction = models.ForeignKey(
        InventoryTransaction,
        null=True,
        blank=True,
        on_delete=models.CASCADE,
        related_name="allocations",
    )
    opening_purchase_accrual = models.ForeignKey(
        OpeningPurchaseAccrual,
        null=True,
        blank=True,
        on_delete=models.CASCADE,
        related_name="allocations",
    )
    opening_goods_issued = models.ForeignKey(
        OpeningGoodsIssued,
        null=True,
        blank=True,
        on_delete=models.CASCADE,
        related_name="allocations",
    )
    contract_line = models.ForeignKey(
        PurchaseContractLine, null=True, blank=True, on_delete=models.CASCADE
    )
    purchase_invoice_line = models.ForeignKey(
        PurchaseInvoiceLine, null=True, blank=True, on_delete=models.CASCADE
    )
    sales_invoice_line = models.ForeignKey(
        SalesInvoiceLine, null=True, blank=True, on_delete=models.CASCADE
    )
    production_order = models.ForeignKey(
        ProductionOrder, null=True, blank=True, on_delete=models.CASCADE
    )
    matched_quantity = models.DecimalField(max_digits=20, decimal_places=6)
    unit_price = models.DecimalField(
        max_digits=20, decimal_places=8, null=True, blank=True
    )
    amount = models.DecimalField(
        max_digits=20, decimal_places=2, null=True, blank=True
    )
    confidence = models.PositiveSmallIntegerField(default=100)
    is_manual = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["kind", "inventory_transaction_id", "id"]


class SettlementRun(models.Model):
    class Status(models.TextChoices):
        RUNNING = "running", "计算中"
        SUCCESS = "success", "成功"
        FAILED = "failed", "失败"

    period = models.ForeignKey(
        AccountingPeriod, on_delete=models.CASCADE, related_name="runs"
    )
    run_id = models.CharField(max_length=36, unique=True)
    status = models.CharField(max_length=20, choices=Status.choices)
    iterations = models.PositiveIntegerField(default=0)
    converged = models.BooleanField(default=False)
    started_at = models.DateTimeField(auto_now_add=True)
    finished_at = models.DateTimeField(null=True, blank=True)
    summary = models.JSONField(default=dict, blank=True)
    error_message = models.TextField(blank=True)

    class Meta:
        ordering = ["-started_at"]


class InventoryValuation(models.Model):
    period = models.ForeignKey(
        AccountingPeriod, on_delete=models.CASCADE, related_name="inventory_valuations"
    )
    material_code = models.CharField(max_length=80)
    material_name = models.CharField(max_length=255, blank=True)
    opening_quantity = models.DecimalField(max_digits=20, decimal_places=6, default=ZERO)
    opening_amount = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    receipt_quantity = models.DecimalField(max_digits=20, decimal_places=6, default=ZERO)
    receipt_amount = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    issue_quantity = models.DecimalField(max_digits=20, decimal_places=6, default=ZERO)
    issue_amount = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    weighted_average_price = models.DecimalField(
        max_digits=20, decimal_places=8, default=ZERO
    )
    ending_quantity = models.DecimalField(max_digits=20, decimal_places=6, default=ZERO)
    ending_amount = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    iterations = models.PositiveIntegerField(default=0)
    converged = models.BooleanField(default=False)

    class Meta:
        ordering = ["material_code"]
        constraints = [
            models.UniqueConstraint(
                fields=["period", "material_code"], name="unique_valuation_per_material"
            )
        ]


class OrderCostResult(models.Model):
    period = models.ForeignKey(
        AccountingPeriod, on_delete=models.CASCADE, related_name="order_cost_results"
    )
    order = models.OneToOneField(
        ProductionOrder, on_delete=models.CASCADE, related_name="cost_result"
    )
    direct_material = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    water_power = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    steam = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    natural_gas = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    depreciation = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    consumables = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    wages = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    total_cost = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    finished_material = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    finished_overhead = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    finished_wages = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    finished_cost = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    wip_material = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    wip_overhead = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    wip_wages = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    wip_cost = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    finished_unit_cost = models.DecimalField(
        max_digits=20, decimal_places=8, default=ZERO
    )

    class Meta:
        ordering = ["order__workshop", "order__order_no"]


class ReconciliationResult(models.Model):
    class Status(models.TextChoices):
        PASS = "pass", "通过"
        WARNING = "warning", "警告"
        FAIL = "fail", "不通过"

    period = models.ForeignKey(
        AccountingPeriod, on_delete=models.CASCADE, related_name="reconciliations"
    )
    check_code = models.CharField(max_length=50)
    check_name = models.CharField(max_length=255)
    left_amount = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    right_amount = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    difference = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    status = models.CharField(max_length=20, choices=Status.choices)
    details = models.JSONField(default=dict, blank=True)

    class Meta:
        ordering = ["check_code"]
        constraints = [
            models.UniqueConstraint(
                fields=["period", "check_code"], name="unique_reconciliation_check"
            )
        ]


class ExceptionItem(models.Model):
    class Severity(models.TextChoices):
        BLOCKING = "blocking", "阻断"
        WARNING = "warning", "警告"
        INFO = "info", "提示"

    class Status(models.TextChoices):
        OPEN = "open", "待处理"
        RESOLVED = "resolved", "已处理"
        IGNORED = "ignored", "已批准忽略"

    period = models.ForeignKey(
        AccountingPeriod, on_delete=models.CASCADE, related_name="exceptions"
    )
    category = models.CharField(max_length=80)
    severity = models.CharField(max_length=20, choices=Severity.choices)
    reference = models.CharField(max_length=255, blank=True)
    message = models.TextField()
    status = models.CharField(
        max_length=20, choices=Status.choices, default=Status.OPEN
    )
    created_at = models.DateTimeField(auto_now_add=True)
    resolved_at = models.DateTimeField(null=True, blank=True)
    resolved_by = models.ForeignKey(
        settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL
    )

    class Meta:
        ordering = ["severity", "category", "id"]
