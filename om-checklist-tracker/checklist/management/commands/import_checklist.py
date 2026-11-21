from django.core.management.base import BaseCommand, CommandError

from checklist.services.workbook_import import import_workbook


class Command(BaseCommand):
    help = "Import checklist definitions from an Excel template."

    def add_arguments(self, parser):
        parser.add_argument("workbook", nargs="?", default="Trafo_Checklist_Nov_2026_Populated.xlsx")
        parser.add_argument("--site", default="Talf Energy India Pvt Ltd, Jhatikara DL9")

    def handle(self, *args, **options):
        try:
            template, count = import_workbook(options["workbook"], options["site"])
        except (OSError, KeyError, ValueError) as exc:
            raise CommandError(f"Could not import workbook: {exc}") from exc
        self.stdout.write(self.style.SUCCESS(f"Imported {count} items into {template} (total: {template.items.count()})."))
