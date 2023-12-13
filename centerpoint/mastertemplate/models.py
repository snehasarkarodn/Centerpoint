from django.db import models

class ProcessedData(models.Model):
    unique_id = models.CharField(max_length=8)
    output_path = models.CharField(max_length=255)
    selected_values = models.JSONField()