# Generated by Django 5.0 on 2023-12-28 13:22

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('mastertemplate', '0003_processeddata_created_by_processeddata_filename'),
    ]

    operations = [
        migrations.AddField(
            model_name='processeddata',
            name='duration_of_creation',
            field=models.DurationField(blank=True, null=True),
        ),
    ]