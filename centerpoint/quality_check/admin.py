from django.contrib import admin
from .models import QualityCheckRecord

class QualityCheckRecordAdmin(admin.ModelAdmin):
    list_display = ('unique_id' ,'file_path' ,'file_name' ,'num_records' ,
                    'date_of_processing' ,'qc_processing_time','qc_done_by')
admin.site.register(QualityCheckRecord,QualityCheckRecordAdmin)


    