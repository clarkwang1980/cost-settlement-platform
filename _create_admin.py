from django.contrib.auth import get_user_model
import os

User = get_user_model()
username = os.getenv("DJANGO_SUPERUSER_USERNAME", "labon")
password = os.getenv("DJANGO_SUPERUSER_PASSWORD", "labon123")

user, created = User.objects.get_or_create(username=username)
user.set_password(password)
user.is_staff = True
user.is_superuser = True
user.save()
print(f"OK: {username} {'CREATED' if created else 'UPDATED'}")
