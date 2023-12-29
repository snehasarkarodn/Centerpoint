from django.db import models

class ProcessedData(models.Model):
    unique_id = models.CharField(max_length=8)
    output_path = models.CharField(max_length=255)
    selected_values = models.JSONField()

    filename = models.CharField(max_length=255, null=True, blank=True)
    created_by = models.CharField(max_length=255, null=True, blank=True)
    created_on = models.DateTimeField(auto_now_add=True)
    duration_of_creation = models.IntegerField(null=True, blank=True)

class SheetUpdate(models.Model):
    file_version = models.CharField(max_length=255)
    edited_by = models.CharField(max_length=255)
    last_edit_date = models.DateTimeField(auto_now=True)
    duration_of_update = models.IntegerField(null=True, blank=True)
    file_path = models.CharField(max_length=255)

    def __str__(self):
        return f"{self.file_version} - {self.edited_by}"
