from __future__ import annotations

import csv
import hashlib
import io
import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from django.db import transaction
from openpyxl import load_workbook

from settlement.models import (
    AccountingPeriod,
    ImportBatch,
    InventoryTransaction,
    InventoryValuation,
    Material,
    MatchAllocation,
    OpeningGoodsIssued,
    OpeningPurchaseAccrual,
    OpeningWip,
    OrderCostResult,
    ProductionOrder,
    PurchaseContractLine,
    PurchaseInvoiceLine,
    ReconciliationResult,
    SalesInvoiceLine,
    SettlementRun,
    WorkshopSettlement,
    WorkshopCostPool,
)


HEADER_CLEANER = re.compile(r"[\s\r\n\t（）()【】\[\]：:、,，._\-]+")


def normalize_header(value) -> str:
    return HEADER_CLEANER.sub("", str(value or "").strip().lower())


ALIASES = {
    "inventory": {
        "txn_date": ["业务日期", "单据日期", "日期", "出入库日期", "记账日期"],
        "document_no": ["单据号", "单据编号", "出入库单号", "凭证号"],
        "line_no": ["行号", "序号", "明细行"],
        "business_type": ["业务类型", "出入库类型", "单据类型", "事务类型"],
        "direction": ["方向", "出入库方向", "收发标志"],
        "warehouse": ["仓库", "仓库名称", "库位"],
        "material_code": ["物料编码", "物料代码", "存货编码", "产品编码", "品号"],
        "material_name": ["物料名称", "存货名称", "产品名称", "品名"],
        "specification": ["规格", "规格型号", "型号"],
        "quantity": ["数量", "实收数量", "实发数量", "出入库数量", "基本数量"],
        "uom": ["单位", "计量单位", "基本单位"],
        "order_no": ["工单号", "生产订单号", "生产订单", "订单号"],
        "workshop": ["车间", "生产车间", "成本中心"],
        "supplier_code": ["供应商编码", "供应商代码"],
        "supplier_name": ["供应商", "供应商名称"],
        "customer_code": ["客户编码", "客户代码"],
        "customer_name": ["客户", "客户名称"],
        "contract_no": ["合同号", "采购合同号", "合同编号"],
        "invoice_no": ["发票号", "发票号码"],
        "unit_price": ["不含税单价", "单价", "入库单价", "单位成本"],
        "amount": ["不含税金额", "金额", "成本金额", "入库金额"],
        "remark": ["备注", "摘要"],
    },
    "purchase_contract": {
        "contract_no": ["合同号", "采购合同号", "合同编号"],
        "contract_date": ["合同日期", "签订日期", "日期"],
        "supplier_code": ["供应商编码", "供应商代码"],
        "supplier_name": ["供应商", "供应商名称"],
        "material_code": ["物料编码", "物料代码", "存货编码"],
        "material_name": ["物料名称", "存货名称", "品名"],
        "quantity": ["合同数量", "数量"],
        "unit_price_excl_tax": ["不含税单价", "合同不含税价", "单价"],
    },
    "purchase_invoice": {
        "invoice_no": ["发票号", "发票号码"],
        "invoice_date": ["发票日期", "开票日期", "日期"],
        "invoice_period": ["发票期间", "开票期间", "所属期间"],
        "contract_no": ["合同号", "采购合同号", "合同编号"],
        "supplier_code": ["供应商编码", "供应商代码"],
        "supplier_name": ["供应商", "供应商名称"],
        "material_code": ["物料编码", "物料代码", "存货编码"],
        "material_name": ["物料名称", "存货名称", "品名"],
        "quantity": ["发票数量", "数量"],
        "unit_price_excl_tax": ["不含税单价", "发票不含税单价", "单价"],
        "amount_excl_tax": ["不含税金额", "金额"],
    },
    "sales_invoice": {
        "invoice_no": ["发票号", "发票号码"],
        "invoice_date": ["发票日期", "开票日期", "日期"],
        "invoice_period": ["发票期间", "开票期间", "所属期间"],
        "sales_order_no": ["销售订单号", "订单号", "销售合同号"],
        "customer_code": ["客户编码", "客户代码"],
        "customer_name": ["客户", "客户名称"],
        "material_code": ["物料编码", "物料代码", "存货编码", "产品编码"],
        "material_name": ["物料名称", "存货名称", "产品名称", "品名"],
        "quantity": ["开票数量", "发票数量", "数量"],
        "unit_price_excl_tax": ["不含税单价", "销售单价", "单价"],
        "amount_excl_tax": ["不含税金额", "金额"],
    },
    "production_order": {
        "order_no": ["工单号", "生产订单号", "生产订单", "订单号"],
        "workshop": ["车间", "生产车间", "成本中心"],
        "product_code": ["产品编码", "产出物料编码", "主产品编码", "物料编码"],
        "product_name": ["产品名称", "产出物料名称", "物料名称"],
        "plan_quantity": ["计划数量", "订单数量", "计划产量"],
        "finished_quantity": ["完工数量", "报工数量", "入库数量"],
        "wip_quantity": ["在产品数量", "在制数量", "未完工数量"],
        "status": ["状态", "完工状态", "订单状态"],
        "standard_hours": ["标准工时", "单位工时", "工时"],
        "material_completion_rate": ["材料完工比例", "材料完成率", "完工比例"],
        "wip_conversion_rate": ["在制转换比例", "在制人工比例", "在制完工程度"],
        "production_receipt_document": ["生产入库单号", "完工入库单号", "入库单号"],
    },
    "workshop_cost": {
        "workshop": ["车间", "生产车间", "成本中心"],
        "water_power": ["水电", "水电费", "电费"],
        "steam": ["蒸汽", "蒸汽费"],
        "natural_gas": ["天然气", "天然气费"],
        "depreciation": ["折旧", "折旧费"],
        "consumables": ["机物料", "机物料消耗", "辅料"],
        "wages": ["工资", "人工", "直接人工"],
    },
    "opening_wip": {
        "order_no": ["工单号", "生产订单号", "订单号"],
        "material": ["材料", "材料成本", "期初材料"],
        "water_power": ["水电", "水电费"],
        "steam": ["蒸汽", "蒸汽费"],
        "natural_gas": ["天然气", "天然气费"],
        "depreciation": ["折旧", "折旧费"],
        "consumables": ["机物料", "机物料消耗"],
        "wages": ["工资", "人工"],
    },
    "opening_purchase_accrual": {
        "original_period": ["原入库月份", "入库月份", "所属月份"],
        "document_date": ["原入库日期", "入库日期", "单据日期"],
        "document_no": ["原入库单号", "入库单号", "单据号"],
        "line_no": ["行号", "序号", "明细行"],
        "supplier_code": ["供应商编码", "供应商代码"],
        "supplier_name": ["供应商", "供应商名称"],
        "contract_no": ["合同号", "采购合同号"],
        "material_code": ["物料编码", "存货编码"],
        "material_name": ["物料名称", "存货名称"],
        "uom": ["单位", "计量单位"],
        "original_quantity": ["原入库数量", "累计入库数量"],
        "prior_invoiced_quantity": [
            "以前累计到票数量",
            "以前到票数量",
            "期前到票数量",
        ],
        "opening_quantity": [
            "月初未到票数量",
            "期初暂估数量",
            "期初未到票数量",
        ],
        "provisional_unit_price": [
            "暂估单价",
            "入库暂估单价",
            "不含税暂估单价",
        ],
        "opening_amount": [
            "月初暂估金额",
            "期初暂估金额",
            "未到票暂估金额",
        ],
        "remark": ["备注", "摘要"],
    },
    "opening_goods_issued": {
        "original_period": ["原出库月份", "出库月份", "所属月份"],
        "document_date": ["原出库日期", "出库日期", "单据日期"],
        "document_no": ["原出库单号", "出库单号", "单据号"],
        "line_no": ["行号", "序号", "明细行"],
        "sales_order_no": ["销售订单号", "订单号", "销售合同号"],
        "customer_code": ["客户编码", "客户代码"],
        "customer_name": ["客户", "客户名称"],
        "material_code": ["产品编码", "物料编码", "存货编码"],
        "material_name": ["产品名称", "物料名称", "存货名称"],
        "uom": ["单位", "计量单位"],
        "original_quantity": ["原出库数量", "累计出库数量"],
        "prior_invoiced_quantity": [
            "以前累计开票数量",
            "以前开票数量",
            "期前开票数量",
        ],
        "opening_quantity": [
            "月初未开票数量",
            "期初发出商品数量",
            "期初未开票数量",
        ],
        "unit_cost": [
            "原出库单位成本",
            "发出商品单位成本",
            "单位成本",
        ],
        "opening_amount": [
            "月初发出商品成本",
            "期初发出商品成本",
            "未开票成本",
        ],
        "remark": ["备注", "摘要"],
    },
}


REQUIRED_FIELDS = {
    "inventory": {"txn_date", "business_type", "material_code", "quantity"},
    "purchase_contract": {
        "contract_no",
        "material_code",
        "quantity",
        "unit_price_excl_tax",
    },
    "purchase_invoice": {
        "invoice_no",
        "invoice_date",
        "material_code",
        "quantity",
        "unit_price_excl_tax",
    },
    "sales_invoice": {
        "invoice_no",
        "invoice_date",
        "material_code",
        "quantity",
    },
    "production_order": {
        "order_no",
        "workshop",
        "product_code",
        "finished_quantity",
    },
    "workshop_cost": {"workshop"},
    "opening_wip": {"order_no"},
    "opening_purchase_accrual": {
        "original_period",
        "document_date",
        "document_no",
        "material_code",
        "opening_quantity",
        "provisional_unit_price",
    },
    "opening_goods_issued": {
        "original_period",
        "document_date",
        "document_no",
        "material_code",
        "opening_quantity",
        "unit_cost",
    },
}


def sha256_uploaded(uploaded_file) -> str:
    digest = hashlib.sha256()
    for chunk in uploaded_file.chunks():
        digest.update(chunk)
    uploaded_file.seek(0)
    return digest.hexdigest()


def to_decimal(value, default=None) -> Decimal | None:
    if value is None or value == "":
        return default
    if isinstance(value, Decimal):
        return value
    if isinstance(value, bool):
        return Decimal(int(value))
    text = str(value).strip().replace(",", "").replace("￥", "").replace("¥", "")
    if not text:
        return default
    is_percent = text.endswith("%")
    if is_percent:
        text = text[:-1]
    try:
        result = Decimal(text)
        return result / Decimal("100") if is_percent else result
    except InvalidOperation as exc:
        raise ValueError(f"无法识别数字：{value}") from exc


def to_date(value, default=None) -> date | None:
    if value is None or value == "":
        return default
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        return date(1899, 12, 30) + __import__("datetime").timedelta(days=int(value))
    text = str(value).strip()
    for fmt in (
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%Y.%m.%d",
        "%Y%m%d",
        "%Y-%m",
        "%Y/%m",
        "%Y%m",
    ):
        try:
            parsed = datetime.strptime(text, fmt).date()
            return parsed.replace(day=1) if "%d" not in fmt else parsed
        except ValueError:
            continue
    raise ValueError(f"无法识别日期：{value}")


def month_start(value, fallback: date) -> date:
    parsed = to_date(value, fallback)
    return parsed.replace(day=1)


def text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _read_csv(path: Path):
    raw = path.read_bytes()
    decoded = None
    for encoding in ("utf-8-sig", "gb18030", "utf-8"):
        try:
            decoded = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if decoded is None:
        raise ValueError("CSV文件编码无法识别，请使用UTF-8或GB18030。")
    return list(csv.reader(io.StringIO(decoded)))


def _read_xlsx(path: Path):
    workbook = load_workbook(path, read_only=True, data_only=True)
    best_rows = None
    best_score = -1
    for worksheet in workbook.worksheets:
        rows = [list(row) for row in worksheet.iter_rows(values_only=True)]
        score = sum(1 for row in rows[:10] for value in row if value not in (None, ""))
        if score > best_score:
            best_score = score
            best_rows = rows
    workbook.close()
    return best_rows or []


def _locate_header(rows, kind: str):
    alias_map = {}
    for field, names in ALIASES[kind].items():
        for name in names:
            alias_map[normalize_header(name)] = field

    best = None
    for index, row in enumerate(rows[:10]):
        mapping = {}
        for col_index, value in enumerate(row):
            field = alias_map.get(normalize_header(value))
            if field and field not in mapping:
                mapping[field] = col_index
        required_hits = len(REQUIRED_FIELDS[kind] & set(mapping))
        score = required_hits * 100 + len(mapping)
        if best is None or score > best[0]:
            best = (score, index, mapping)

    if not best:
        raise ValueError("找不到表头。")
    _, header_index, mapping = best
    missing = REQUIRED_FIELDS[kind] - set(mapping)
    if missing:
        labels = "、".join(sorted(missing))
        raise ValueError(f"缺少必填字段：{labels}")
    return header_index, mapping


def read_rows(path: Path, kind: str):
    suffix = path.suffix.lower()
    if suffix == ".csv":
        rows = _read_csv(path)
    elif suffix in {".xlsx", ".xlsm"}:
        rows = _read_xlsx(path)
    else:
        raise ValueError("只支持CSV、XLSX和XLSM文件。")
    if not rows:
        raise ValueError("文件没有可读取的数据。")
    header_index, mapping = _locate_header(rows, kind)
    result = []
    for source_row_number, row in enumerate(rows[header_index + 1 :], header_index + 2):
        if not any(value not in (None, "") for value in row):
            continue
        item = {"_row_number": source_row_number}
        for field, col_index in mapping.items():
            item[field] = row[col_index] if col_index < len(row) else None
        result.append(item)
    return result


def infer_direction(value, business_type: str) -> str:
    normalized = text(value).lower()
    if normalized in {"入", "入库", "in", "收入"}:
        return InventoryTransaction.Direction.IN
    if normalized in {"出", "出库", "out", "发出"}:
        return InventoryTransaction.Direction.OUT
    kind = business_type.lower()
    if "退料" in kind or "入库" in kind or "盘盈" in kind:
        return InventoryTransaction.Direction.IN
    if any(token in kind for token in ("出库", "领料", "发货", "盘亏", "采购退货")):
        return InventoryTransaction.Direction.OUT
    raise ValueError(f"无法根据业务类型判断出入库方向：{business_type}")


def normalize_order_status(value) -> str:
    normalized = text(value).lower()
    if normalized in {"完工", "完成", "已完工", "closed", "completed", "1", "是"}:
        return ProductionOrder.Status.COMPLETED
    return ProductionOrder.Status.OPEN


def get_or_create_material(code, name="", specification="", uom=""):
    material, _ = Material.objects.get_or_create(
        code=text(code),
        defaults={
            "name": text(name),
            "specification": text(specification),
            "uom": text(uom),
        },
    )
    changed = False
    if not material.name and name:
        material.name = text(name)
        changed = True
    if not material.specification and specification:
        material.specification = text(specification)
        changed = True
    if not material.uom and uom:
        material.uom = text(uom)
        changed = True
    if changed:
        material.save(update_fields=["name", "specification", "uom"])
    return material


def clear_calculation_outputs(period: AccountingPeriod):
    MatchAllocation.objects.filter(period=period).delete()
    InventoryValuation.objects.filter(period=period).delete()
    OrderCostResult.objects.filter(period=period).delete()
    ReconciliationResult.objects.filter(period=period).delete()
    SettlementRun.objects.filter(period=period).delete()
    period.exceptions.all().delete()
    WorkshopSettlement.objects.filter(period=period).update(
        status=WorkshopSettlement.Status.PENDING,
        iterations=0,
        converged=False,
        settled_at=None,
        settled_by=None,
        order_cost=Decimal("0"),
        production_receipt_amount=Decimal("0"),
        error_message="",
    )
    period.status = AccountingPeriod.Status.OPEN
    period.last_run_at = None
    period.save(update_fields=["status", "last_run_at"])


def _delete_existing_kind(period: AccountingPeriod, kind: str):
    ImportBatch.objects.filter(period=period, kind=kind).delete()


def create_import_batch(*, period, kind, uploaded_file, user, replace_existing=False):
    if period.status == AccountingPeriod.Status.LOCKED:
        raise ValueError("该期间已锁账，不能继续导入。")
    if (
        kind
        in {
            ImportBatch.Kind.OPENING_PURCHASE_ACCRUAL,
            ImportBatch.Kind.OPENING_GOODS_ISSUED,
        }
        and AccountingPeriod.objects.filter(month__lt=period.month).exists()
    ):
        raise ValueError(
            "期初采购暂估和期初发出商品仅用于系统首个核算期间；"
            "后续月份由历史明细自动滚动。"
        )
    file_hash = sha256_uploaded(uploaded_file)
    if ImportBatch.objects.filter(period=period, kind=kind, sha256=file_hash).exists():
        raise ValueError("相同文件已经导入过，无需重复上传。")
    if replace_existing:
        _delete_existing_kind(period, kind)
    batch = ImportBatch.objects.create(
        period=period,
        kind=kind,
        file=uploaded_file,
        original_name=uploaded_file.name,
        sha256=file_hash,
        uploaded_by=user,
    )
    return batch


@transaction.atomic
def process_batch(batch: ImportBatch):
    rows = read_rows(Path(batch.file.path), batch.kind)
    clear_calculation_outputs(batch.period)
    importer = {
        ImportBatch.Kind.INVENTORY: _import_inventory,
        ImportBatch.Kind.PURCHASE_CONTRACT: _import_purchase_contract,
        ImportBatch.Kind.PURCHASE_INVOICE: _import_purchase_invoice,
        ImportBatch.Kind.SALES_INVOICE: _import_sales_invoice,
        ImportBatch.Kind.PRODUCTION_ORDER: _import_production_order,
        ImportBatch.Kind.WORKSHOP_COST: _import_workshop_cost,
        ImportBatch.Kind.OPENING_WIP: _import_opening_wip,
        ImportBatch.Kind.OPENING_PURCHASE_ACCRUAL: (
            _import_opening_purchase_accrual
        ),
        ImportBatch.Kind.OPENING_GOODS_ISSUED: _import_opening_goods_issued,
    }[batch.kind]
    errors = []
    count = 0
    for row in rows:
        try:
            importer(batch, row)
            count += 1
        except Exception as exc:
            errors.append(f"第{row['_row_number']}行：{exc}")
    batch.row_count = count
    batch.error_count = len(errors)
    batch.message = "\n".join(errors[:100])
    batch.status = (
        ImportBatch.Status.SUCCESS if count > 0 else ImportBatch.Status.FAILED
    )
    batch.save(update_fields=["row_count", "error_count", "message", "status"])
    if count == 0:
        raise ValueError(batch.message or "没有成功导入任何数据。")
    return batch


def _import_inventory(batch, row):
    business_type = text(row.get("business_type"))
    direction = infer_direction(row.get("direction"), business_type)
    quantity = abs(to_decimal(row.get("quantity"), Decimal("0")))
    if quantity == 0:
        raise ValueError("数量不能为0")
    material_code = text(row.get("material_code"))
    material = get_or_create_material(
        material_code,
        row.get("material_name"),
        row.get("specification"),
        row.get("uom"),
    )
    unit_price = to_decimal(row.get("unit_price"))
    amount = to_decimal(row.get("amount"))
    if amount is None and unit_price is not None:
        amount = (quantity * unit_price).quantize(Decimal("0.01"))
    InventoryTransaction.objects.create(
        period=batch.period,
        batch=batch,
        txn_date=to_date(row.get("txn_date")),
        document_no=text(row.get("document_no")),
        line_no=text(row.get("line_no") or row["_row_number"]),
        business_type=business_type,
        direction=direction,
        warehouse=text(row.get("warehouse")),
        material=material,
        material_code=material_code,
        material_name=text(row.get("material_name")),
        specification=text(row.get("specification")),
        quantity=quantity,
        uom=text(row.get("uom")),
        order_no=text(row.get("order_no")),
        workshop=text(row.get("workshop")),
        supplier_code=text(row.get("supplier_code")),
        supplier_name=text(row.get("supplier_name")),
        customer_code=text(row.get("customer_code")),
        customer_name=text(row.get("customer_name")),
        contract_no=text(row.get("contract_no")),
        invoice_no=text(row.get("invoice_no")),
        unit_price=unit_price,
        amount=amount,
        price_source="原始金额" if amount is not None else "",
        match_status=InventoryTransaction.MatchStatus.UNMATCHED,
        remark=text(row.get("remark")),
    )


def _import_purchase_contract(batch, row):
    quantity = abs(to_decimal(row.get("quantity"), Decimal("0")))
    PurchaseContractLine.objects.create(
        period=batch.period,
        batch=batch,
        contract_no=text(row.get("contract_no")),
        contract_date=to_date(row.get("contract_date")),
        supplier_code=text(row.get("supplier_code")),
        supplier_name=text(row.get("supplier_name")),
        material_code=text(row.get("material_code")),
        material_name=text(row.get("material_name")),
        quantity=quantity,
        remaining_quantity=quantity,
        unit_price_excl_tax=to_decimal(row.get("unit_price_excl_tax"), Decimal("0")),
    )


def _import_purchase_invoice(batch, row):
    quantity = abs(to_decimal(row.get("quantity"), Decimal("0")))
    unit_price = to_decimal(row.get("unit_price_excl_tax"), Decimal("0"))
    amount = to_decimal(row.get("amount_excl_tax"))
    if amount is None:
        amount = (quantity * unit_price).quantize(Decimal("0.01"))
    invoice_date = to_date(row.get("invoice_date"))
    PurchaseInvoiceLine.objects.create(
        period=batch.period,
        batch=batch,
        invoice_no=text(row.get("invoice_no")),
        invoice_date=invoice_date,
        invoice_period=month_start(row.get("invoice_period"), invoice_date),
        contract_no=text(row.get("contract_no")),
        supplier_code=text(row.get("supplier_code")),
        supplier_name=text(row.get("supplier_name")),
        material_code=text(row.get("material_code")),
        material_name=text(row.get("material_name")),
        quantity=quantity,
        unit_price_excl_tax=unit_price,
        amount_excl_tax=amount,
        remaining_quantity=quantity,
    )


def _import_sales_invoice(batch, row):
    quantity = abs(to_decimal(row.get("quantity"), Decimal("0")))
    unit_price = to_decimal(row.get("unit_price_excl_tax"))
    amount = to_decimal(row.get("amount_excl_tax"))
    if amount is None and unit_price is not None:
        amount = (quantity * unit_price).quantize(Decimal("0.01"))
    invoice_date = to_date(row.get("invoice_date"))
    SalesInvoiceLine.objects.create(
        period=batch.period,
        batch=batch,
        invoice_no=text(row.get("invoice_no")),
        invoice_date=invoice_date,
        invoice_period=month_start(row.get("invoice_period"), invoice_date),
        sales_order_no=text(row.get("sales_order_no")),
        customer_code=text(row.get("customer_code")),
        customer_name=text(row.get("customer_name")),
        material_code=text(row.get("material_code")),
        material_name=text(row.get("material_name")),
        quantity=quantity,
        unit_price_excl_tax=unit_price,
        amount_excl_tax=amount,
        remaining_quantity=quantity,
    )


def _import_production_order(batch, row):
    plan_quantity = abs(to_decimal(row.get("plan_quantity"), Decimal("0")))
    finished_quantity = abs(to_decimal(row.get("finished_quantity"), Decimal("0")))
    wip_quantity = to_decimal(row.get("wip_quantity"))
    if wip_quantity is None:
        wip_quantity = max(plan_quantity - finished_quantity, Decimal("0"))
    status = normalize_order_status(row.get("status"))
    if status == ProductionOrder.Status.COMPLETED:
        wip_quantity = Decimal("0")
    ProductionOrder.objects.update_or_create(
        period=batch.period,
        order_no=text(row.get("order_no")),
        defaults={
            "batch": batch,
            "carried_from": None,
            "is_system_generated": False,
            "workshop": text(row.get("workshop")),
            "product_code": text(row.get("product_code")),
            "product_name": text(row.get("product_name")),
            "plan_quantity": plan_quantity,
            "finished_quantity": finished_quantity,
            "wip_quantity": abs(wip_quantity),
            "status": status,
            "standard_hours": to_decimal(
                row.get("standard_hours"), Decimal("1")
            ),
            "material_completion_rate": to_decimal(
                row.get("material_completion_rate")
            ),
            "wip_conversion_rate": to_decimal(
                row.get("wip_conversion_rate"), Decimal("0.5")
            ),
            "production_receipt_document": text(
                row.get("production_receipt_document")
            ),
        },
    )


def _import_workshop_cost(batch, row):
    WorkshopCostPool.objects.create(
        period=batch.period,
        batch=batch,
        workshop=text(row.get("workshop")),
        water_power=to_decimal(row.get("water_power"), Decimal("0")),
        steam=to_decimal(row.get("steam"), Decimal("0")),
        natural_gas=to_decimal(row.get("natural_gas"), Decimal("0")),
        depreciation=to_decimal(row.get("depreciation"), Decimal("0")),
        consumables=to_decimal(row.get("consumables"), Decimal("0")),
        wages=to_decimal(row.get("wages"), Decimal("0")),
    )


def _import_opening_wip(batch, row):
    OpeningWip.objects.update_or_create(
        period=batch.period,
        order_no=text(row.get("order_no")),
        defaults={
            "batch": batch,
            "source_period": None,
            "is_system_generated": False,
            "material": to_decimal(row.get("material"), Decimal("0")),
            "water_power": to_decimal(row.get("water_power"), Decimal("0")),
            "steam": to_decimal(row.get("steam"), Decimal("0")),
            "natural_gas": to_decimal(
                row.get("natural_gas"), Decimal("0")
            ),
            "depreciation": to_decimal(
                row.get("depreciation"), Decimal("0")
            ),
            "consumables": to_decimal(
                row.get("consumables"), Decimal("0")
            ),
            "wages": to_decimal(row.get("wages"), Decimal("0")),
        },
    )


def _opening_ledger_values(batch, row, *, price_field, ledger_name):
    original_period = month_start(row.get("original_period"), batch.period.month)
    document_date = to_date(row.get("document_date"))
    if original_period >= batch.period.month:
        raise ValueError(f"{ledger_name}的原业务月份必须早于系统首个核算月份。")
    if document_date is None or document_date >= batch.period.month:
        raise ValueError(f"{ledger_name}的原单据日期必须早于系统首个核算月份。")
    if document_date.replace(day=1) != original_period:
        raise ValueError("原业务月份必须与原单据日期所在月份一致。")

    opening_quantity = to_decimal(row.get("opening_quantity"))
    prior_quantity = to_decimal(
        row.get("prior_invoiced_quantity"), Decimal("0")
    )
    if opening_quantity is None or opening_quantity <= 0:
        raise ValueError("期初未核销数量必须大于0。")
    if prior_quantity is None or prior_quantity < 0:
        raise ValueError("以前累计核销数量不能小于0。")

    original_quantity = to_decimal(row.get("original_quantity"))
    if original_quantity is None:
        original_quantity = opening_quantity + prior_quantity
    if original_quantity <= 0 or prior_quantity > original_quantity:
        raise ValueError("原业务数量或以前累计核销数量不合理。")
    expected_opening = original_quantity - prior_quantity
    if abs(expected_opening - opening_quantity) > Decimal("0.000001"):
        raise ValueError(
            "期初未核销数量应等于原业务数量减以前累计核销数量。"
        )

    unit_price = to_decimal(row.get(price_field))
    if unit_price is None or unit_price < 0:
        raise ValueError("期初单价或单位成本不能小于0。")
    opening_amount = to_decimal(row.get("opening_amount"))
    if opening_amount is None:
        opening_amount = (opening_quantity * unit_price).quantize(
            Decimal("0.01")
        )
    if opening_amount < 0:
        raise ValueError("期初金额不能小于0。")

    get_or_create_material(
        row.get("material_code"),
        row.get("material_name"),
        uom=row.get("uom"),
    )
    return {
        "original_period": original_period,
        "document_date": document_date,
        "document_no": text(row.get("document_no")),
        "line_no": text(row.get("line_no")) or str(row["_row_number"]),
        "material_code": text(row.get("material_code")),
        "material_name": text(row.get("material_name")),
        "uom": text(row.get("uom")),
        "original_quantity": original_quantity,
        "prior_invoiced_quantity": prior_quantity,
        "opening_quantity": opening_quantity,
        "opening_amount": opening_amount,
        "remaining_quantity": opening_quantity,
        "remark": text(row.get("remark")),
    }, unit_price


def _import_opening_purchase_accrual(batch, row):
    values, unit_price = _opening_ledger_values(
        batch,
        row,
        price_field="provisional_unit_price",
        ledger_name="期初采购暂估",
    )
    OpeningPurchaseAccrual.objects.create(
        period=batch.period,
        batch=batch,
        supplier_code=text(row.get("supplier_code")),
        supplier_name=text(row.get("supplier_name")),
        contract_no=text(row.get("contract_no")),
        provisional_unit_price=unit_price,
        **values,
    )


def _import_opening_goods_issued(batch, row):
    values, unit_cost = _opening_ledger_values(
        batch,
        row,
        price_field="unit_cost",
        ledger_name="期初发出商品",
    )
    OpeningGoodsIssued.objects.create(
        period=batch.period,
        batch=batch,
        sales_order_no=text(row.get("sales_order_no")),
        customer_code=text(row.get("customer_code")),
        customer_name=text(row.get("customer_name")),
        unit_cost=unit_cost,
        **values,
    )
