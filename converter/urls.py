from django.urls import path
from .views import FileUploadView

urlpatterns = [
    path('tables/', FileUploadView.as_view(), name='file-upload'),
]
