from django.contrib.auth import get_user_model
from django.core.management.base import BaseCommand, CommandError


class Command(BaseCommand):
    help = "Create the initial administrator account."

    def add_arguments(self, parser):
        parser.add_argument("username")
        parser.add_argument("--email", default="")
        parser.add_argument("--password", help="Omit to enter a password securely at the prompt.")

    def handle(self, *args, **options):
        from getpass import getpass

        User = get_user_model()
        username = options["username"]
        if User.objects.filter(username=username).exists():
            raise CommandError(f"User {username!r} already exists.")
        password = options["password"] or getpass("Admin password: ")
        if not password:
            raise CommandError("A password is required.")
        User.objects.create_superuser(username=username, email=options["email"], password=password)
        self.stdout.write(self.style.SUCCESS(f"Administrator {username!r} created."))
