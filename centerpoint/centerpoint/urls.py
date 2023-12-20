from django.contrib import admin
from django.urls import include, path

urlpatterns = [
    path('admin/', admin.site.urls),
    path('mastertemplate/', include('mastertemplate.urls')),
    path('quality_check/', include('quality_check.urls')),
]
