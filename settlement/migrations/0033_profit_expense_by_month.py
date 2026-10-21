# -*- coding: utf-8 -*-
"""损益表费用：从「每期间一条」改为「每期间 + 每月一条」（按月录入、各月自己摊）。"""
import django.db.models.deletion
from django.conf import settings
from django.db import migrations, models


def fill_month(apps, schema_editor):
    """存量记录（旧口径：每个期间只有一条）按期间所属月份补齐 month。"""
    Expense = apps.get_model("settlement", "ProfitStatementExpense")
    for expense in Expense.objects.all():
        if expense.month is None:
            expense.month = expense.period.month
            expense.save(update_fields=["month"])


class Migration(migrations.Migration):

    dependencies = [
        ("settlement", "0032_profitstatementexpense"),
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
    ]

    operations = [
        migrations.AlterModelOptions(
            name="profitstatementexpense",
            options={
                "ordering": ["month"],
                "verbose_name": "损益表费用录入",
                "verbose_name_plural": "损益表费用录入",
            },
        ),
        migrations.AddField(
            model_name="profitstatementexpense",
            name="month",
            field=models.DateField(
                null=True, verbose_name="费用所属月份(月首，按出库月份)"
            ),
        ),
        migrations.RunPython(fill_month, migrations.RunPython.noop),
        migrations.AlterField(
            model_name="profitstatementexpense",
            name="month",
            field=models.DateField(verbose_name="费用所属月份(月首，按出库月份)"),
        ),
        migrations.AlterField(
            model_name="profitstatementexpense",
            name="period",
            field=models.ForeignKey(
                on_delete=django.db.models.deletion.CASCADE,
                related_name="profit_expenses",
                to="settlement.accountingperiod",
                verbose_name="核算期间",
            ),
        ),
        migrations.AddConstraint(
            model_name="profitstatementexpense",
            constraint=models.UniqueConstraint(
                fields=("period", "month"), name="unique_profit_expense_period_month"
            ),
        ),
    ]
