from django.urls import path
from .views import FileUploadView

urlpatterns = [
    path('tables/', FileUploadView.as_view(), name='file-upload'),
    path('file/', FileUploadView.as_view(), name='convert-file'),
]
