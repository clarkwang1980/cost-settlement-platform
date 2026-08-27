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
        OPENING_INVENTORY = (
            "opening_inventory",
            "期初存货",
        )
        ENDING_INVENTORY = (
            "ending_inventory",
            "期末存货",
        )
        WORK_ORDER = (
            "work_order",
            "生产工单",
        )
        MATERIAL_ISSUE = (
            "material_issue",
            "投料记录",
        )
        WORK_REPORT = (
            "work_report",
            "报工记录",
        )
        PULP_RECEIPT = (
            "pulp_receipt",
            "浆粕成品入库",
        )
        BOM = (
            "bom",
            "BOM表",
        )
        ISSUE_PROGRESS = (
            "issue_progress",
            "投料进度表",
        )
        HISTORICAL_PRICE = (
            "historical_price",
            "历史单价",
        )
        WORK_ORDER_DETAIL = (
            "work_order_detail",
            "工单明细",
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
    category_name = models.CharField(max_length=255, blank=True, verbose_name="分类名称")
    batch_no = models.CharField(max_length=100, blank=True, verbose_name="批号")
    quantity = models.DecimalField(max_digits=20, decimal_places=6)
    auxiliary_quantity = models.DecimalField(
        max_digits=20,
        decimal_places=6,
        null=True,
        blank=True,
        verbose_name="辅助数量",
    )
    uom = models.CharField(max_length=30, blank=True)
    auxiliary_uom = models.CharField(max_length=30, blank=True, verbose_name="辅助单位")
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
    uom = models.CharField(max_length=50, blank=True)
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


class WorkOrder(models.Model):
    """生产工单主档（ERP 导入台账）。

    按用户给定的 ERP 导出格式存储，不参与现有成本核算流程。
    """

    period = models.ForeignKey(
        AccountingPeriod, on_delete=models.CASCADE, related_name="work_orders"
    )
    batch = models.ForeignKey(
        ImportBatch,
        null=True,
        blank=True,
        on_delete=models.CASCADE,
        related_name="work_orders",
    )
    order_no = models.CharField(max_length=100, db_index=True, verbose_name="工单编号")
    order_type = models.CharField(max_length=100, blank=True, verbose_name="工单类型")
    identifier = models.CharField(max_length=200, blank=True, verbose_name="标识码")
    plan_quantity = models.DecimalField(
        max_digits=20, decimal_places=6, default=ZERO, verbose_name="计划数量"
    )
    planned_start = models.DateTimeField(null=True, blank=True, verbose_name="计划开始时间")
    planned_finish = models.DateTimeField(null=True, blank=True, verbose_name="计划完工时间")
    priority = models.CharField(max_length=50, blank=True, verbose_name="优先级")
    line_equipment = models.CharField(max_length=200, blank=True, verbose_name="产线设备")
    shift = models.CharField(max_length=100, blank=True, verbose_name="班次")
    production_dept_code = models.CharField(
        max_length=100, blank=True, verbose_name="生产部门编号"
    )
    production_supervisor_code = models.CharField(
        max_length=100, blank=True, verbose_name="生产主管编号"
    )
    planning_dept_code = models.CharField(
        max_length=100, blank=True, verbose_name="计划部门编号"
    )
    planner_code = models.CharField(max_length=100, blank=True, verbose_name="计划员编号")
    remark = models.CharField(max_length=500, blank=True, verbose_name="备注")
    specified_material = models.CharField(
        max_length=255, blank=True, verbose_name="指定用料"
    )
    external_document_no = models.CharField(
        max_length=100, blank=True, verbose_name="外部单据编号"
    )
    status = models.CharField(max_length=50, blank=True, verbose_name="状态")
    use_bom = models.BooleanField(default=False, verbose_name="是否使用物料清单")
    use_routing = models.BooleanField(default=False, verbose_name="是否使用工艺路线")
    process_plan_type = models.CharField(
        max_length=100, blank=True, verbose_name="工序计划类型"
    )
    enable_sop = models.BooleanField(default=False, verbose_name="启用SOP")
    production_mode = models.CharField(max_length=100, blank=True, verbose_name="生产模式")
    line_code = models.CharField(max_length=100, blank=True, verbose_name="产线编号")
    source_order_no = models.CharField(max_length=100, blank=True, verbose_name="订单号")
    material_color = models.CharField(max_length=100, blank=True, verbose_name="物料颜色")
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["order_no"]
        constraints = [
            models.UniqueConstraint(
                fields=["period", "order_no"], name="unique_work_order_per_period"
            )
        ]
        verbose_name = "生产工单"
        verbose_name_plural = "生产工单"

    def __str__(self):
        return self.order_no


class MaterialIssue(models.Model):
    """投料记录（ERP 导入台账）。

    按用户给定的 ERP 导出格式存储，不参与现有成本核算流程。
    """

    period = models.ForeignKey(
        AccountingPeriod, on_delete=models.CASCADE, related_name="material_issues"
    )
    batch = models.ForeignKey(
        ImportBatch,
        null=True,
        blank=True,
        on_delete=models.CASCADE,
        related_name="material_issues",
    )
    task_no = models.CharField(max_length=100, blank=True, db_index=True, verbose_name="生产任务编号")
    operation_no = models.CharField(max_length=100, blank=True, verbose_name="工序编号")
    operation_name = models.CharField(max_length=200, blank=True, verbose_name="工序名称")
    material_code = models.CharField(max_length=80, db_index=True, verbose_name="物料编码")
    material_name = models.CharField(max_length=255, blank=True, verbose_name="物料名称")
    material_category = models.CharField(max_length=255, blank=True, verbose_name="物料分类")
    material_attribute = models.CharField(max_length=255, blank=True, verbose_name="物料属性")
    material_spec = models.CharField(max_length=255, blank=True, verbose_name="物料规格")
    issue_quantity = models.DecimalField(
        max_digits=20, decimal_places=6, default=ZERO, verbose_name="投料数量"
    )
    returned_quantity = models.DecimalField(
        max_digits=20, decimal_places=6, default=ZERO, verbose_name="已回撤数量"
    )
    actual_issue_quantity = models.DecimalField(
        max_digits=20, decimal_places=6, default=ZERO, verbose_name="实际投料数量"
    )
    issue_uom = models.CharField(max_length=30, blank=True, verbose_name="投料单位")
    auxiliary_quantity = models.DecimalField(
        max_digits=20,
        decimal_places=6,
        null=True,
        blank=True,
        verbose_name="辅助单位数量",
    )
    auxiliary_uom = models.CharField(max_length=30, blank=True, verbose_name="辅助单位")
    identifier = models.CharField(max_length=200, blank=True, verbose_name="标识码")
    batch_no = models.CharField(max_length=100, blank=True, verbose_name="批次号")
    quality_status = models.CharField(max_length=50, blank=True, verbose_name="质量状态")
    issue_location = models.CharField(max_length=100, blank=True, verbose_name="投料仓位")
    replaced_material_code = models.CharField(
        max_length=80, blank=True, verbose_name="被替代物料编号"
    )
    replaced_material_name = models.CharField(
        max_length=255, blank=True, verbose_name="被替代物料名称"
    )
    issue_time = models.DateTimeField(null=True, blank=True, verbose_name="投料时间")
    issuer = models.CharField(max_length=100, blank=True, verbose_name="投料人员")
    executor = models.CharField(max_length=100, blank=True, verbose_name="执行人员")
    work_order_no = models.CharField(max_length=100, db_index=True, verbose_name="工单编号")
    main_output_code = models.CharField(
        max_length=80, blank=True, verbose_name="主产出物料编号"
    )
    main_output_name = models.CharField(
        max_length=255, blank=True, verbose_name="主产出物料名称"
    )
    sales_order_no = models.CharField(max_length=100, blank=True, verbose_name="销售订单编号")
    record_no = models.CharField(max_length=100, blank=True, verbose_name="投料记录编号")
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["work_order_no", "material_code", "id"]
        verbose_name = "投料记录"
        verbose_name_plural = "投料记录"

    def __str__(self):
        return f"{self.work_order_no} {self.material_code}"


class WorkReport(models.Model):
    """报工记录（ERP 导入台账）。

    按用户给定的 ERP 导出格式存储，不参与现有成本核算流程。
    """

    period = models.ForeignKey(
        AccountingPeriod, on_delete=models.CASCADE, related_name="work_reports"
    )
    batch = models.ForeignKey(
        ImportBatch,
        null=True,
        blank=True,
        on_delete=models.CASCADE,
        related_name="work_reports",
    )
    report_time = models.DateTimeField(null=True, blank=True, verbose_name="报工时间")
    task_no = models.CharField(max_length=100, blank=True, db_index=True, verbose_name="生产任务编号")
    operation_no = models.CharField(max_length=100, blank=True, verbose_name="工序编号")
    operation_name = models.CharField(max_length=200, blank=True, verbose_name="工序名称")
    material_code = models.CharField(max_length=80, db_index=True, verbose_name="报工物料编号")
    material_name = models.CharField(max_length=255, blank=True, verbose_name="报工物料名称")
    operator = models.CharField(max_length=100, blank=True, verbose_name="生产人员")
    quality_status = models.CharField(max_length=50, blank=True, verbose_name="质量状态")
    defect_reason = models.CharField(max_length=500, blank=True, verbose_name="不良原因")
    quantity = models.DecimalField(
        max_digits=20, decimal_places=6, default=ZERO, verbose_name="报工数量"
    )
    uom = models.CharField(max_length=30, blank=True, verbose_name="报工单位")
    auxiliary_quantity = models.DecimalField(
        max_digits=20,
        decimal_places=6,
        null=True,
        blank=True,
        verbose_name="辅助数量",
    )
    auxiliary_uom = models.CharField(max_length=30, blank=True, verbose_name="辅助单位")
    batch_no = models.CharField(max_length=100, blank=True, verbose_name="批次号")
    identifier = models.CharField(max_length=200, blank=True, verbose_name="标识码")
    route_card_no = models.CharField(max_length=100, blank=True, verbose_name="流转卡编号")
    work_hours = models.DecimalField(
        max_digits=20, decimal_places=6, null=True, blank=True, verbose_name="工时"
    )
    work_hours_uom = models.CharField(max_length=30, blank=True, verbose_name="工时单位")
    remark = models.CharField(max_length=500, blank=True, verbose_name="备注")
    department = models.CharField(max_length=100, blank=True, verbose_name="部门")
    equipment = models.CharField(max_length=200, blank=True, verbose_name="设备")
    work_order_no = models.CharField(max_length=100, db_index=True, verbose_name="工单编号")
    work_order_material_code = models.CharField(
        max_length=80, blank=True, verbose_name="工单物料编号"
    )
    work_order_material_name = models.CharField(
        max_length=255, blank=True, verbose_name="工单物料名称"
    )
    order_no = models.CharField(max_length=100, blank=True, verbose_name="订单编号")
    paper_tube_color = models.CharField(max_length=100, blank=True, verbose_name="特纺/包覆-纸管颜色")
    spinning_shift = models.CharField(max_length=100, blank=True, verbose_name="纺纱-班别")
    moisture_content = models.CharField(max_length=100, blank=True, verbose_name="染色纤维-含水率")
    dye_batch_flow = models.CharField(max_length=100, blank=True, verbose_name="染色纤维-批号流水")
    weaving_sequence = models.CharField(max_length=100, blank=True, verbose_name="织布-序号")
    weaving_machine = models.CharField(max_length=100, blank=True, verbose_name="织布-机台")
    output_location = models.CharField(max_length=100, blank=True, verbose_name="产出仓位")
    work_order_department = models.CharField(
        max_length=100, blank=True, verbose_name="工单部门"
    )
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["work_order_no", "report_time", "id"]
        verbose_name = "报工记录"
        verbose_name_plural = "报工记录"

    def __str__(self):
        return f"{self.work_order_no} {self.material_code}"


class BomLine(models.Model):
    """BOM 表：按工单展开的子项物料清单（用于中间品自动流转与投料校验）。"""

    period = models.ForeignKey(
        AccountingPeriod, on_delete=models.CASCADE, related_name="bom_lines"
    )
    batch = models.ForeignKey(
        ImportBatch, on_delete=models.CASCADE, related_name="bom_lines"
    )
    work_order_no = models.CharField(max_length=100, db_index=True)
    material_code = models.CharField(max_length=80, blank=True, verbose_name="成品物料编号")
    material_name = models.CharField(max_length=255, blank=True, verbose_name="成品物料名称")
    line_no = models.CharField(max_length=40, blank=True, verbose_name="项次")
    sub_item_code = models.CharField(max_length=80, db_index=True, verbose_name="子项物料编号")
    sub_item_name = models.CharField(max_length=255, blank=True, verbose_name="子项物料名称")
    qty_numerator = models.DecimalField(
        max_digits=20, decimal_places=6, default=1, verbose_name="子项用量分子"
    )
    qty_denominator = models.DecimalField(
        max_digits=20, decimal_places=6, default=1, verbose_name="子项用量分母"
    )
    uom = models.CharField(max_length=30, blank=True, verbose_name="单位")
    unit_price = models.DecimalField(
        max_digits=20, decimal_places=6, null=True, blank=True, verbose_name="子项单价"
    )
    standard_price = models.DecimalField(
        max_digits=20, decimal_places=6, null=True, blank=True, verbose_name="标准成品单价"
    )

    class Meta:
        ordering = ["work_order_no", "line_no", "id"]
        verbose_name = "BOM表"
        verbose_name_plural = "BOM表"

    def __str__(self):
        return f"{self.work_order_no} {self.sub_item_code}"

    @property
    def qty(self):
        """子项用量（分子/分母）。"""
        denominator = self.qty_denominator or 1
        if not denominator:
            return 1
        return self.qty_numerator / denominator


class IssueProgress(models.Model):
    """工单投料进度表（ERP 导出台账）：投料状态等用于工单成本页展示。"""

    period = models.ForeignKey(
        AccountingPeriod, on_delete=models.CASCADE, related_name="issue_progress"
    )
    batch = models.ForeignKey(
        ImportBatch, on_delete=models.CASCADE, related_name="issue_progress"
    )
    create_date = models.CharField(max_length=50, blank=True, verbose_name="创建日期")
    department = models.CharField(max_length=100, blank=True, verbose_name="部门名称")
    work_order_no = models.CharField(max_length=100, db_index=True, verbose_name="工单编号")
    line_no = models.CharField(max_length=40, blank=True, verbose_name="行号")
    material_code = models.CharField(max_length=80, db_index=True, verbose_name="物料编号")
    apply_quantity = models.DecimalField(
        max_digits=20, decimal_places=6, default=ZERO, verbose_name="申请数量"
    )
    issue_quantity = models.DecimalField(
        max_digits=20, decimal_places=6, default=ZERO, verbose_name="发料数量"
    )
    issue_batch_mark = models.CharField(
        max_length=100, blank=True, verbose_name="发料批次-标识码"
    )
    actual_issue_quantity = models.DecimalField(
        max_digits=20, decimal_places=6, default=ZERO, verbose_name="投料数量"
    )
    actual_batch_mark = models.CharField(
        max_length=100, blank=True, verbose_name="投料批次-标识码"
    )
    uom = models.CharField(max_length=30, blank=True, verbose_name="单位")
    issue_status = models.CharField(max_length=50, blank=True, verbose_name="投料状态")

    class Meta:
        ordering = ["work_order_no", "line_no", "id"]
        verbose_name = "投料进度表"
        verbose_name_plural = "投料进度表"

    def __str__(self):
        return f"{self.work_order_no} {self.material_code} {self.issue_status}"


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
    dye_auxiliary = models.DecimalField(
        max_digits=20, decimal_places=2, default=ZERO, verbose_name="染料助剂费"
    )
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
    quantity = models.DecimalField(
        max_digits=20, decimal_places=6, default=ZERO, verbose_name="期初在产品数量"
    )
    material = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    water_power = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    steam = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    natural_gas = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    depreciation = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    consumables = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    dye_auxiliary = models.DecimalField(
        max_digits=20, decimal_places=2, default=ZERO, verbose_name="染料助剂费"
    )
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
            + self.dye_auxiliary
            + self.wages
        )


class OpeningInventoryBalance(models.Model):
    period = models.ForeignKey(
        AccountingPeriod,
        on_delete=models.CASCADE,
        related_name="opening_inventory_balances",
    )
    batch = models.ForeignKey(
        ImportBatch,
        null=True,
        blank=True,
        on_delete=models.CASCADE,
        related_name="opening_inventory_balances",
    )
    source_period = models.ForeignKey(
        AccountingPeriod,
        null=True,
        blank=True,
        on_delete=models.PROTECT,
        related_name="carried_inventory_balances",
    )
    material_code = models.CharField(max_length=80, db_index=True)
    material_name = models.CharField(max_length=255, blank=True)
    opening_date = models.DateField(null=True, blank=True, verbose_name="期初日期")
    category_name = models.CharField(max_length=255, blank=True, verbose_name="分类名称")
    warehouse = models.CharField(
        max_length=100, blank=True, null=True, verbose_name="仓库名称"
    )
    batch_no = models.CharField(
        max_length=100, blank=True, null=True, verbose_name="批次号"
    )
    uom = models.CharField(max_length=30, blank=True, verbose_name="单位名称")
    quantity = models.DecimalField(max_digits=20, decimal_places=6, default=ZERO)
    amount = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    unit_price = models.DecimalField(max_digits=20, decimal_places=8, default=ZERO)
    is_system_generated = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["material_code"]
        constraints = [
            models.UniqueConstraint(
                fields=["period", "material_code", "warehouse", "batch_no"],
                name="unique_opening_inventory_batch",
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


class EndingInventoryBalance(models.Model):
    period = models.ForeignKey(
        AccountingPeriod,
        on_delete=models.CASCADE,
        related_name="ending_inventory_balances",
    )
    batch = models.ForeignKey(
        ImportBatch,
        null=True,
        blank=True,
        on_delete=models.CASCADE,
        related_name="ending_inventory_balances",
    )
    material_code = models.CharField(max_length=80, db_index=True)
    material_name = models.CharField(max_length=255, blank=True)
    ending_date = models.DateField(null=True, blank=True, verbose_name="期末日期")
    category_name = models.CharField(max_length=255, blank=True, verbose_name="分类名称")
    warehouse = models.CharField(
        max_length=100, blank=True, null=True, verbose_name="仓库名称"
    )
    batch_no = models.CharField(
        max_length=100, blank=True, null=True, verbose_name="批次号"
    )
    uom = models.CharField(max_length=30, blank=True, verbose_name="单位名称")
    quantity = models.DecimalField(max_digits=20, decimal_places=6, default=ZERO)
    amount = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    unit_price = models.DecimalField(max_digits=20, decimal_places=8, default=ZERO)
    is_system_generated = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["material_code"]
        constraints = [
            models.UniqueConstraint(
                fields=["period", "material_code", "warehouse", "batch_no"],
                name="unique_ending_inventory_batch",
            )
        ]

    def __str__(self):
        return f"{self.period} {self.material_code} {self.quantity}"


class MatchAllocation(models.Model):
    class Kind(models.TextChoices):
        PURCHASE_CONTRACT = "purchase_contract", "采购入库-合同"
        PURCHASE_INVOICE = "purchase_invoice", "采购入库-发票"
        SALES_INVOICE = "sales_invoice", "销售出库-发票"

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


class HistoricalPrice(models.Model):
    """历史单价（用户上传）：期间+物料 的入库单价/出库单价，用于单价表优先展示。

    上传后覆盖单价表中该期间该物料的系统计算单价；未上传的物料仍用系统月末加权价。
    """

    period = models.ForeignKey(
        AccountingPeriod, on_delete=models.CASCADE, related_name="historical_prices"
    )
    material_code = models.CharField(max_length=80, db_index=True)
    material_name = models.CharField(max_length=255, blank=True)
    receipt_price = models.DecimalField(
        max_digits=20, decimal_places=8, null=True, blank=True,
        verbose_name="入库单价",
    )
    issue_price = models.DecimalField(
        max_digits=20, decimal_places=8, null=True, blank=True,
        verbose_name="出库单价",
    )

    class Meta:
        ordering = ["material_code"]
        constraints = [
            models.UniqueConstraint(
                fields=["period", "material_code"],
                name="unique_historical_price_per_period_material",
            )
        ]
        verbose_name = "历史单价"
        verbose_name_plural = "历史单价"

    def __str__(self):
        return f"{self.period} {self.material_code}"


class WorkOrderDetail(models.Model):
    """工单明细（用户上传）：创建时间/车间/工单号/计划完工/实际完工/完成状态/计划/投料/入库数量。

    按工单编号保存用户上传的明细，导出中心"工单明细"报表按上传原样导出。
    """

    period = models.ForeignKey(
        AccountingPeriod, on_delete=models.CASCADE, related_name="work_order_details"
    )
    batch = models.ForeignKey(
        ImportBatch, on_delete=models.CASCADE, related_name="work_order_details"
    )
    created_at = models.CharField(max_length=100, blank=True, verbose_name="创建时间")
    workshop = models.CharField(max_length=100, blank=True, verbose_name="生产车间")
    order_no = models.CharField(max_length=100, db_index=True, verbose_name="工单编号")
    category = models.CharField(max_length=100, blank=True, verbose_name="类别")
    planned_finish = models.CharField(max_length=100, blank=True, verbose_name="计划完工时间")
    actual_finish = models.CharField(max_length=100, blank=True, verbose_name="实际完工时间")
    completion_status = models.CharField(max_length=100, blank=True, verbose_name="完成状态")
    plan_quantity = models.DecimalField(max_digits=20, decimal_places=6, default=ZERO)
    issue_quantity = models.DecimalField(max_digits=20, decimal_places=6, default=ZERO)
    receipt_quantity = models.DecimalField(max_digits=20, decimal_places=6, default=ZERO)
    row_no = models.PositiveIntegerField(default=0, verbose_name="上传行号")

    class Meta:
        ordering = ["row_no", "id"]
        constraints = [
            models.UniqueConstraint(
                fields=["period", "order_no"], name="unique_work_order_detail_per_period"
            )
        ]
        verbose_name = "工单明细"
        verbose_name_plural = "工单明细"

    def __str__(self):
        return f"{self.period} {self.order_no}"


class WorkOrderCostResult(models.Model):
    """工单成本结果（按生产工单+投料记录+报工记录核算）。"""

    period = models.ForeignKey(
        AccountingPeriod,
        on_delete=models.CASCADE,
        related_name="work_order_cost_results",
    )
    work_order = models.OneToOneField(
        WorkOrder, on_delete=models.CASCADE, related_name="cost_result"
    )
    product_code = models.CharField(max_length=80, blank=True, verbose_name="产品编码")
    product_name = models.CharField(max_length=255, blank=True, verbose_name="产品名称")
    plan_quantity = models.DecimalField(max_digits=20, decimal_places=6, default=ZERO)
    finished_quantity = models.DecimalField(max_digits=20, decimal_places=6, default=ZERO)
    wip_quantity = models.DecimalField(max_digits=20, decimal_places=6, default=ZERO)
    direct_material = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    water_power = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    steam = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    natural_gas = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    depreciation = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    consumables = models.DecimalField(max_digits=20, decimal_places=2, default=ZERO)
    dye_auxiliary = models.DecimalField(
        max_digits=20, decimal_places=2, default=ZERO, verbose_name="染料助剂费"
    )
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
        ordering = ["work_order__production_dept_code", "work_order__order_no"]
        verbose_name = "工单成本"
        verbose_name_plural = "工单成本"

    def __str__(self):
        return self.work_order.order_no


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
