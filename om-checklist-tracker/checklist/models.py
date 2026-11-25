from datetime import time

from django.conf import settings
from django.core.exceptions import ValidationError
from django.db import models
from django.utils import timezone


class Site(models.Model):
    name = models.CharField(max_length=160, unique=True)
    company = models.CharField(max_length=200, blank=True)
    location = models.CharField(max_length=200, blank=True)
    active = models.BooleanField(default=True)

    def __str__(self):
        return self.name


class SiteEquipment(models.Model):
    site = models.ForeignKey(Site, on_delete=models.CASCADE, related_name="equipment")
    name = models.CharField(max_length=200)

    class Meta:
        ordering = ["name"]
        constraints = [models.UniqueConstraint(fields=["site", "name"], name="unique_equipment_per_site")]

    def __str__(self):
        return self.name


class UserSite(models.Model):
    user = models.OneToOneField(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name="site_assignment")
    site = models.OneToOneField(Site, on_delete=models.CASCADE, related_name="user_assignment")

    def __str__(self):
        return f"{self.user} → {self.site}"


class ChecklistTemplate(models.Model):
    name = models.CharField(max_length=160)
    document_number = models.CharField(max_length=120, blank=True)
    company = models.CharField(max_length=200, blank=True)
    equipment_name = models.CharField(max_length=200, blank=True)
    version = models.PositiveIntegerField(default=1)
    active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["name", "-version"]
        constraints = [models.UniqueConstraint(fields=["name", "version"], name="unique_template_version")]

    def __str__(self):
        return f"{self.name} v{self.version}"


class ChecklistItem(models.Model):
    class AnswerType(models.TextChoices):
        NUMBER = "number", "Numeric reading"
        BOOLEAN = "boolean", "Yes / No"
        TEXT = "text", "Short value / text"
        CHOICES = "choices", "Dropdown choices"

    template = models.ForeignKey(ChecklistTemplate, on_delete=models.CASCADE, related_name="items")
    section = models.CharField(max_length=160, blank=True)
    checkpoint_no = models.PositiveIntegerField()
    description = models.TextField()
    standard = models.TextField(blank=True)
    method = models.CharField(max_length=120, blank=True)
    frequency = models.CharField(max_length=80, blank=True)
    answer_type = models.CharField(max_length=12, choices=AnswerType.choices, default=AnswerType.TEXT)
    # Comma-separated list of allowed option values when answer_type == "choices".
    # E.g. "Operational,Non Operational,Spare"
    choices = models.CharField(max_length=500, blank=True)
    unit = models.CharField(max_length=40, blank=True)
    required = models.BooleanField(default=True)
    position = models.PositiveIntegerField(default=0)

    def get_choices_list(self):
        """Return a list of option strings for a 'choices' type item."""
        if self.answer_type == self.AnswerType.CHOICES and self.choices:
            return [c.strip() for c in self.choices.split(",") if c.strip()]
        return []

    class Meta:
        ordering = ["position", "checkpoint_no"]
        constraints = [models.UniqueConstraint(fields=["template", "checkpoint_no"], name="unique_checkpoint_per_template")]

    def __str__(self):
        return f"{self.checkpoint_no}. {self.description[:70]}"


class SiteChecklist(models.Model):
    site = models.ForeignKey(Site, on_delete=models.CASCADE, related_name="checklists")
    template = models.ForeignKey(ChecklistTemplate, on_delete=models.CASCADE, related_name="site_checklists")
    position = models.PositiveIntegerField(default=0)
    active = models.BooleanField(default=True)

    class Meta:
        ordering = ["position", "template__name"]
        constraints = [models.UniqueConstraint(fields=["site", "template"], name="unique_template_per_site")]

    def __str__(self):
        return f"{self.site} — {self.template}"


class Submission(models.Model):
    site = models.ForeignKey(Site, on_delete=models.PROTECT, related_name="submissions")
    template = models.ForeignKey(ChecklistTemplate, on_delete=models.PROTECT, related_name="submissions")
    submitted_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.PROTECT, related_name="checklist_submissions")
    checklist_date = models.DateField()
    submitted_at = models.DateTimeField(default=timezone.now)
    remarks = models.TextField(blank=True)
    signer_name = models.CharField(max_length=160, blank=True)
    equipment_snapshot = models.TextField(blank=True)

    class Meta:
        ordering = ["-checklist_date", "site__name"]
        constraints = [models.UniqueConstraint(fields=["site", "template", "checklist_date"], name="one_submission_per_site_template_per_day")]

    def clean(self):
        if self.submitted_by_id and not self.submitted_by.is_superuser:
            assignment = getattr(self.submitted_by, "site_assignment", None)
            if assignment is None or assignment.site_id != self.site_id:
                raise ValidationError("This user is not assigned to the selected site.")
        if self.submitted_by_id:
            assignment = getattr(self.submitted_by, "site_assignment", None)
            if not self.submitted_by.is_superuser and assignment and UserSite.objects.filter(site_id=assignment.site_id).exclude(user_id=self.submitted_by_id).exists():
                raise ValidationError("This site is assigned to a different user.")

    def __str__(self):
        return f"{self.site} — {self.checklist_date}"


class Answer(models.Model):
    submission = models.ForeignKey(Submission, on_delete=models.CASCADE, related_name="answers")
    item = models.ForeignKey(ChecklistItem, on_delete=models.PROTECT, related_name="answers")
    value = models.CharField(max_length=300, blank=True)
    comment = models.TextField(blank=True)

    class Meta:
        constraints = [models.UniqueConstraint(fields=["submission", "item"], name="one_answer_per_item")]

    def __str__(self):
        return f"{self.item.checkpoint_no}: {self.value}"


class GlobalSettings(models.Model):
    deadline = models.TimeField(default=time(17, 0))
    updated_at = models.DateTimeField(auto_now=True)
    updated_by = models.ForeignKey(settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL)

    def save(self, *args, **kwargs):
        self.pk = 1
        super().save(*args, **kwargs)

    @classmethod
    def get_solo(cls):
        obj, _ = cls.objects.get_or_create(pk=1)
        return obj

    @classmethod
    def is_open(cls, now=None):
        now = timezone.localtime(now) if now is not None else timezone.localtime()
        return now.time().replace(tzinfo=None) < cls.get_solo().deadline

    def __str__(self):
        return f"Daily deadline: {self.deadline:%H:%M} IST"
