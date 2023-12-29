from django.urls import path
from .views import process_file, download_template

urlpatterns = [
    path('process_file', process_file, name='process_file'),
    path('download/<path:file_path>/', download_template, name='download_template'),
]
