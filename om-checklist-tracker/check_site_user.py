import django
import os

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "config.settings")
django.setup()

from checklist.models import Site, UserSite, Submission

for s in Site.objects.all():
    user = getattr(s, "user_assignment", None)
    print(f"Site: {s.name}, User: {user.user.username if user else 'None'}")
