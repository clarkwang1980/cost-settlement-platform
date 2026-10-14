from settlement.models import CustomerOrder, CustomerReceipt
from django.contrib import admin

from settlement import models


@admin.register(models.AccountingPeriod)
class AccountingPeriodAdmin(admin.ModelAdmin):
    list_display = ("month", "status", "last_run_at", "locked_at", "locked_by")
    list_filter = ("status",)


@admin.register(models.InventoryTransaction)
class InventoryTransactionAdmin(admin.ModelAdmin):
    list_display = (
        "txn_date",
        "document_no",
        "business_type",
        "direction",
        "material_code",
        "quantity",
        "unit_price",
        "amount",
        "order_no",
        "match_mode",
        "match_status",
    )
    list_filter = (
        "period",
        "direction",
        "business_type",
        "match_mode",
        "match_status",
    )
    search_fields = (
        "document_no",
        "material_code",
        "material_name",
        "order_no",
        "contract_no",
        "invoice_no",
    )


for model in (
    models.ImportBatch,
    models.Material,
    models.PurchaseContractLine,
    models.PurchaseInvoiceLine,
    models.PulpMaterial,
    models.SalesInvoiceLine,
    models.WorkshopCostPool,
    models.WorkshopSettlement,
    models.OpeningWip,
    models.OpeningInventoryBalance,
    models.OpeningPurchaseAccrual,
    models.OpeningGoodsIssued,
    models.WorkOrder,
    models.MaterialIssue,
    models.WorkReport,
    models.MatchAllocation,
    models.SettlementRun,
    models.InventoryValuation,
    models.WorkOrderCostResult,
    models.ReconciliationResult,
    models.ExceptionItem,
):
    admin.site.register(model)


admin.site.site_header = "月度成本结算平台管理"
admin.site.site_title = "成本结算管理"
admin.site.index_title = "基础数据与规则维护"


@admin.register(CustomerOrder)
class CustomerOrderAdmin(admin.ModelAdmin):
    list_display = ("customer_name", "order_no", "contract_no", "sign_year",
                    "product_name", "product_category", "quantity", "amount", "due_date")
    list_filter = ("product_category", "sign_year")
    search_fields = ("customer_name", "order_no", "contract_no", "product_name")


@admin.register(CustomerReceipt)
class CustomerReceiptAdmin(admin.ModelAdmin):
    list_display = ("customer_name", "receipt_date", "contract_no", "order_no", "amount")
    list_filter = ("customer_name",)
    search_fields = ("customer_name", "contract_no", "order_no")
