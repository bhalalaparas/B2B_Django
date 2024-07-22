from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('upload_files/', views.upload_files, name='upload_files'),
    path('download/<path:file_path>/', views.download_file, name='download_file'),
]
