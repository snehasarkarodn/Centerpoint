from django.db import models

class ProcessedData(models.Model):
    unique_id = models.CharField(max_length=8)
    output_path = models.CharField(max_length=255)
    selected_values = models.JSONField()

    filename = models.CharField(max_length=255, null=True, blank=True)
    created_by = models.CharField(max_length=255, null=True, blank=True)
    created_on = models.DateTimeField(auto_now_add=True)