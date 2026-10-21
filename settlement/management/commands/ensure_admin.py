import os

from django.contrib.auth import get_user_model
from django.core.management.base import BaseCommand


class Command(BaseCommand):
    help = "根据环境变量创建首个管理员账号。"

    def handle(self, *args, **options):
        username = os.getenv("DJANGO_SUPERUSER_USERNAME")
        password = os.getenv("DJANGO_SUPERUSER_PASSWORD")
        email = os.getenv("DJANGO_SUPERUSER_EMAIL", "")
        if not username or not password:
            self.stdout.write("未设置管理员环境变量，跳过自动创建。")
            return
        user_model = get_user_model()
        if user_model.objects.filter(username=username).exists():
            user = user_model.objects.get(username=username)
            user.set_password(password)
            user.is_staff = True
            user.is_superuser = True
            if email:
                user.email = email
            user.save()
            self.stdout.write(self.style.SUCCESS(f"管理员 {username} 密码已更新。"))
            return
        user_model.objects.create_superuser(
            username=username, email=email, password=password
        )
        self.stdout.write(self.style.SUCCESS(f"管理员 {username} 已创建。"))

