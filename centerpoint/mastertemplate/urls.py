from django.urls import path
from .views import main_func, download_template, update_sheet


urlpatterns = [
    path('template', main_func, name='main_func'),
    path('download/<path:file_path>/', download_template, name='download_template'),
    path('update_sheet', update_sheet, name='update_sheet'),
]