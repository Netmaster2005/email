from django.urls import path
from .views import RegisterUser, RetrieveEmails

urlpatterns = [
    path('email/', RegisterUser.as_view()),
    path('email/retrieve', RetrieveEmails.as_view()),
]