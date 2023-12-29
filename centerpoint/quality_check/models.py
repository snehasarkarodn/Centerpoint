from django.db import models

class QualityCheckRecord(models.Model):
    unique_id = models.CharField(max_length=50, primary_key=True)
    file_path = models.CharField(max_length=255)
    file_name = models.CharField(max_length=100)
    num_records = models.IntegerField()
    date_of_processing = models.DateTimeField(auto_now_add=True)
    qc_processing_time = models.IntegerField()
    qc_done_by = models.CharField(max_length=100)

    def __str__(self):
        return f"QualityCheckRecord - {self.unique_id}"
