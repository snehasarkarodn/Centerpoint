from django.urls import path
from .views import main_func, download_template, update_sheet, download_latest_sheet


urlpatterns = [
    path('template', main_func, name='main_func'),
    path('download/<path:file_path>/', download_template, name='download_template'),
    path('update_sheet', update_sheet, name='update_sheet'),
    path('download_latest_sheet', download_latest_sheet, name='download_latest_sheet'),
]