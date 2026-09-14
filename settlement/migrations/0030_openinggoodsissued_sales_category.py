# -*- coding: utf-8 -*-
# Generated manually: add 销售大类(sales_category) to OpeningGoodsIssued
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("settlement", "0029_reconciliationresult_confirmed_and_more"),
    ]

    operations = [
        migrations.AddField(
            model_name="openinggoodsissued",
            name="sales_category",
            field=models.CharField(
                blank=True, max_length=120, verbose_name="销售大类"
            ),
        ),
    ]
