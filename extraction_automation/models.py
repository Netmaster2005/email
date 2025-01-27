from django.db import models
from django.core.validators import EmailValidator

class User(models.Model):
    email = models.EmailField(
        unique=True,
        validators=[EmailValidator(message="Enter a valid email address")],
        error_messages={"unique": "This email is already in use."},
    )
    password = models.CharField(max_length=128)

    def __str__(self):
        return self.email
