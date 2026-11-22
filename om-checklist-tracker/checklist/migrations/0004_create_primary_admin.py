# Neutralized migration replacing legacy hardcoded primary admin creation

from django.db import migrations


def noop(apps, schema_editor):
    pass


class Migration(migrations.Migration):

    dependencies = [
        ("checklist", "0003_add_sitechecklist_and_update_constraint"),
    ]

    operations = [
        migrations.RunPython(noop, noop),
    ]

