from django.contrib import admin
from .models import ProcessedData, SheetUpdate

class ProcessedDataAdmin(admin.ModelAdmin):
    list_display = ('unique_id' ,'output_path','selected_values' ,
                    'filename' ,'created_by' ,'created_on' ,'duration_of_creation')
admin.site.register(ProcessedData,ProcessedDataAdmin)


class SheetUpdateAdmin(admin.ModelAdmin):
    list_display = ('file_version' ,'edited_by' ,'last_edit_date' ,'duration_of_update' ,'file_path')

admin.site.register(SheetUpdate,SheetUpdateAdmin)