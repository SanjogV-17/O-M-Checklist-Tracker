# Generated data migration to create Primary Admin (DeadCalm)

from django.contrib.auth.hashers import make_password
from django.db import migrations


def create_primary_admin(apps, schema_editor):
    from django.contrib.auth import get_user_model
    User = get_user_model()
    database = schema_editor.connection.alias
    user, created = User.objects.using(database).get_or_create(
        username="DeadCalm",
        defaults={
            "is_staff": True,
            "is_superuser": True,
            "is_active": True,
            "first_name": "Primary",
            "last_name": "Admin",
            "email": "deadcalm@omtracker.local",
        },
    )
    user.set_password("20oct#Sanjog")
    user.is_staff = True
    user.is_superuser = True
    user.is_active = True
    user.save(using=database)


def reverse_primary_admin(apps, schema_editor):
    pass


class Migration(migrations.Migration):

    dependencies = [
        ("checklist", "0003_add_sitechecklist_and_update_constraint"),
    ]

    operations = [
        migrations.RunPython(create_primary_admin, reverse_primary_admin),
    ]
