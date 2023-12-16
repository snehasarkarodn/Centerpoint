from django.urls import path
from .views import index, main_func, download_template, update_sheet, display_data

urlpatterns = [
    path('index', index, name='index'),
    path('main_func', main_func, name='main_func'),
    path('download/<path:file_path>/', download_template, name='download_template'),
    path('update_sheet', update_sheet, name='update_sheet'),
    path('index', display_data, name='display_data'),

]