# urls.py
from django.urls import path
from .views import fetch_or_process_emails

urlpatterns = [
    path('fetch-emails/', fetch_or_process_emails, name='fetch_emails'),
]