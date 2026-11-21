import django
import os

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "config.settings")
django.setup()

from checklist.models import ChecklistTemplate, ChecklistItem, SiteChecklist, Site

print("=== TEMPLATES ===")
for t in ChecklistTemplate.objects.all():
    print(f"ID={t.id}, Name='{t.name}', Version={t.version}, Active={t.active}, ItemsCount={t.items.count()}")
    for item in t.items.all()[:5]:
        print(f"   Item #{item.checkpoint_no}: '{item.description}' (sec='{item.section}', type={item.answer_type})")

print("\n=== SITES AND ASSIGNED CHECKLISTS ===")
for s in Site.objects.all():
    cls = list(s.checklists.select_related("template").all())
    print(f"Site '{s.name}': {[c.template.name for c in cls]}")
