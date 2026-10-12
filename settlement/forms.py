from decimal import Decimal

from django import forms

from settlement.models import (
    AccountingPeriod,
    ImportBatch,
    MatchAllocation,
    MatchingMode,
)


class PeriodForm(forms.ModelForm):
    month = forms.DateField(
        label="核算月份",
        widget=forms.DateInput(attrs={"type": "month"}),
        input_formats=["%Y-%m", "%Y-%m-%d"],
    )

    class Meta:
        model = AccountingPeriod
        fields = ["month"]


class UploadForm(forms.Form):
    # 用户 2026-09-30：「车间费用」不再上传（改在平台「车间费用」页录入）→ 从上传下拉里去掉；
    # 模型里的 kind 保留，因为页面录入的数据仍然用这个 kind 存批次记录。
    kind = forms.ChoiceField(label="资料类型", choices=[
        c for c in ImportBatch.Kind.choices
        if c[0] not in ("pulp_flow", "pulp_material", "workshop_cost")
    ])
    file = forms.FileField(
        label="选择文件",
        help_text="支持CSV、XLSX和XLSM；同一文件不会重复导入。",
    )
    replace_existing = forms.BooleanField(
        label="替换该期间同类资料", required=False, initial=False
    )


class ManualMatchForm(forms.Form):
    transaction_id = forms.IntegerField(label="库存流水ID", min_value=1)
    source_id = forms.IntegerField(label="合同/发票/订单明细ID", min_value=1)
    matched_quantity = forms.DecimalField(
        label="匹配数量", min_value=0, max_digits=20, decimal_places=6
    )


class MatchModeForm(forms.Form):
    match_mode = forms.ChoiceField(
        label="处理方式",
        choices=MatchingMode.choices,
    )
    ignore_reason = forms.CharField(
        label="忽略原因",
        required=False,
        max_length=500,
        widget=forms.TextInput(
            attrs={"placeholder": "选择忽略时建议填写原因"}
        ),
    )


class SourceDocumentMatchForm(forms.Form):
    transaction_id = forms.IntegerField(label="库存流水", min_value=1)
    matched_quantity = forms.DecimalField(
        label="匹配数量",
        min_value=Decimal("0.000001"),
        max_digits=20,
        decimal_places=6,
    )


class OpeningDocumentMatchForm(forms.Form):
    opening_id = forms.IntegerField(label="期初台账明细", min_value=1)
    matched_quantity = forms.DecimalField(
        label="匹配数量",
        min_value=Decimal("0.000001"),
        max_digits=20,
        decimal_places=6,
    )


class InventoryDocumentMatchForm(SourceDocumentMatchForm):
    kind = forms.ChoiceField(
        label="匹配类型",
        choices=MatchAllocation.Kind.choices,
    )
    source_id = forms.IntegerField(label="目标单据", min_value=1)
