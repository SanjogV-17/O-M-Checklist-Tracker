from datetime import date, datetime, time
from io import BytesIO
from unittest.mock import patch

from django.contrib.auth import get_user_model
from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import TestCase
from django.utils import timezone
from openpyxl import Workbook, load_workbook

from .models import Answer, ChecklistItem, ChecklistTemplate, GlobalSettings, Site, SiteChecklist, SiteEquipment, Submission, UserSite

User = get_user_model()


class ChecklistWorkflowTests(TestCase):
    def setUp(self):
        self.site = Site.objects.create(name="North Solar")
        SiteEquipment.objects.create(site=self.site, name="Transformer 1")
        self.user = User.objects.create_user(username="operator", password="secure-test-password")
        UserSite.objects.create(user=self.user, site=self.site)
        self.template = ChecklistTemplate.objects.create(name="Test Trafo", version=1)
        self.number_item = ChecklistItem.objects.create(template=self.template, checkpoint_no=1, description="Winding temperature", standard="Below 80 C", answer_type="number", position=1)
        self.boolean_item = ChecklistItem.objects.create(template=self.template, checkpoint_no=2, description="Oil leak check", standard="No leakage", answer_type="boolean", position=2)
        settings = GlobalSettings.get_solo()
        settings.deadline = time(23, 59)
        settings.save()

    def form_data(self, **extra):
        values = {
            f"item_{self.number_item.id}": "62.5",
            f"comment_{self.number_item.id}": "Within expected range",
            f"item_{self.boolean_item.id}": "YES",
            f"comment_{self.boolean_item.id}": "No visible leak",
            "remarks": "Routine inspection completed",
        }
        values.update(extra)
        return values

    def post_preview(self, data=None):
        return self.client.post("/checklist/", data or self.form_data())

    def test_preview_does_not_persist_submission_and_shows_all_values(self):
        self.client.force_login(self.user)
        response = self.post_preview()
        self.assertEqual(response.status_code, 200)
        self.assertTemplateUsed(response, "checklist/checklist_preview.html")
        self.assertContains(response, "62.5")
        self.assertContains(response, "Within expected range")
        self.assertContains(response, "Routine inspection completed")
        self.assertFalse(Submission.objects.exists())
        self.assertFalse(Answer.objects.exists())

    def test_user_can_finalize_after_review_and_read_saved_submission(self):
        self.client.force_login(self.user)
        self.post_preview()
        lockin = self.client.get("/checklist/finalize/")
        self.assertEqual(lockin.status_code, 200)
        self.assertTemplateUsed(lockin, "checklist/checklist_lockin.html")
        response = self.client.post("/checklist/finalize/", {"confirm_lockin": "yes"})
        submission = Submission.objects.get(site=self.site, checklist_date=timezone.localdate())
        self.assertRedirects(response, f"/submissions/{submission.pk}/review/")
        self.assertEqual(submission.answers.count(), 2)
        self.assertEqual(submission.answers.get(item=self.number_item).comment, "Within expected range")
        self.assertEqual(submission.remarks, "Routine inspection completed")
        review = self.client.get(f"/submissions/{submission.pk}/review/")
        self.assertEqual(review.status_code, 200)
        self.assertContains(review, "No visible leak")
        self.assertContains(review, "Routine inspection completed")
        self.assertEqual(submission.equipment_snapshot, "Transformer 1")

    def test_invalid_preview_does_not_create_draft_or_submission(self):
        self.client.force_login(self.user)
        response = self.post_preview(self.form_data(**{f"item_{self.number_item.id}": "not a number"}))
        self.assertEqual(response.status_code, 200)
        self.assertFalse(Submission.objects.exists())
        self.assertFalse(self.client.session.get("checklist_draft_" + str(self.user.pk)))

    def test_cutoff_is_checked_for_preview_and_final_lockin(self):
        settings = GlobalSettings.get_solo()
        settings.deadline = time(17, 0)
        settings.save()
        self.client.force_login(self.user)
        with patch("checklist.models.timezone.localtime", return_value=timezone.make_aware(datetime(2026, 11, 20, 16, 59))):
            self.post_preview()
        with patch("checklist.models.timezone.localtime", return_value=timezone.make_aware(datetime(2026, 11, 20, 16, 59))):
            self.client.get("/checklist/finalize/")
        with patch("checklist.models.timezone.localtime", return_value=timezone.make_aware(datetime(2026, 11, 20, 17, 0))):
            response = self.client.post("/checklist/finalize/", {"confirm_lockin": "yes"})
        self.assertRedirects(response, "/")
        self.assertFalse(Submission.objects.exists())

    def test_finalize_requires_existing_server_side_preview(self):
        self.client.force_login(self.user)
        response = self.client.get("/checklist/finalize/")
        self.assertRedirects(response, "/checklist/")
        self.assertFalse(Submission.objects.exists())

    def test_lockin_requires_explicit_confirmation(self):
        self.client.force_login(self.user)
        self.post_preview()
        response = self.client.post("/checklist/finalize/", {})
        self.assertRedirects(response, "/checklist/finalize/")
        self.assertFalse(Submission.objects.exists())

    def test_regular_user_cannot_review_another_users_submission_or_download_pdf(self):
        other = User.objects.create_user(username="other", password="secure-test-password")
        other_site = Site.objects.create(name="South Solar")
        UserSite.objects.create(user=other, site=other_site)
        submission = Submission.objects.create(site=other_site, template=self.template, submitted_by=other, checklist_date=date(2026, 11, 19))
        self.client.force_login(self.user)
        self.assertEqual(self.client.get(f"/submissions/{submission.pk}/review/").status_code, 403)
        self.assertEqual(self.client.get(f"/admin/submissions/{submission.pk}/pdf/").status_code, 302)

    def test_operator_excel_review_contains_answers_and_comments(self):
        submission = Submission.objects.create(site=self.site, template=self.template, submitted_by=self.user, checklist_date=date(2026, 11, 19), remarks="All good")
        Answer.objects.create(submission=submission, item=self.number_item, value="65", comment="Checked")
        self.client.force_login(self.user)
        response = self.client.get(f"/submissions/{submission.pk}/excel/")
        self.assertEqual(response["Content-Type"], "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        workbook = load_workbook(BytesIO(response.content), data_only=True)
        self.assertIn("Submission review", workbook.sheetnames)
        rows = list(workbook["Submission review"].values)
        self.assertIn((1, "Winding temperature", "65", "Checked"), rows)
        self.assertIn(("Overall remarks", "All good", None, None), rows)

    def test_admin_can_export_pdf_and_excel(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        submission = Submission.objects.create(site=self.site, template=self.template, submitted_by=admin, checklist_date=date(2026, 11, 19))
        Answer.objects.create(submission=submission, item=self.number_item, value="65", comment="Checked")
        self.client.force_login(admin)
        pdf = self.client.get(f"/admin/submissions/{submission.id}/pdf/")
        xlsx = self.client.get(f"/admin/submissions/{submission.id}/xlsx/")
        self.assertEqual(pdf["Content-Type"], "application/pdf")
        self.assertTrue(pdf.content.startswith(b"%PDF"))
        self.assertEqual(xlsx["Content-Type"], "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        workbook = load_workbook(BytesIO(xlsx.content), data_only=True)
        self.assertEqual(workbook["spv"]["Z7"].value, "65")
        self.assertIn("Submission review", workbook.sheetnames)

    def test_admin_can_edit_user_and_assignment(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        second_site = Site.objects.create(name="East Solar")
        self.client.force_login(admin)
        response = self.client.post(f"/manage/users/{self.user.pk}/edit/", {
            "first_name": "Asha", "last_name": "Operator", "email": "asha@example.com",
            "is_active": "on", "site": str(second_site.pk),
        })
        self.assertRedirects(response, "/manage/users/")
        self.user.refresh_from_db()
        self.assertEqual(self.user.first_name, "Asha")
        self.assertEqual(self.user.email, "asha@example.com")
        self.assertEqual(self.user.site_assignment.site, second_site)
        self.assertEqual(self.user.username, "operator")

    def test_admin_cannot_assign_user_to_an_already_assigned_site(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        other = User.objects.create_user(username="other", password="secure-test-password")
        UserSite.objects.create(user=other, site=Site.objects.create(name="West Solar"))
        self.client.force_login(admin)
        response = self.client.post(f"/manage/users/{self.user.pk}/edit/", {
            "first_name": "", "last_name": "", "email": "", "is_active": "on",
            "site": str(other.site_assignment.site_id),
        })
        self.assertEqual(response.status_code, 200)
        self.assertEqual(self.user.site_assignment.site, self.site)

    def test_admin_can_edit_site_and_add_remove_equipment(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        self.client.force_login(admin)
        response = self.client.post(f"/manage/sites/{self.site.pk}/edit/", {"name": "North Solar", "company": "Sun Co", "location": "Rajasthan", "active": "on"})
        self.assertRedirects(response, f"/manage/sites/{self.site.pk}/edit/")
        add_response = self.client.post(f"/manage/sites/{self.site.pk}/equipment/add/", {"name": "Inverter 1"})
        self.assertRedirects(add_response, f"/manage/sites/{self.site.pk}/edit/")
        equipment = SiteEquipment.objects.get(site=self.site, name="Inverter 1")
        remove_response = self.client.post(f"/manage/sites/{self.site.pk}/equipment/{equipment.pk}/remove/")
        self.assertRedirects(remove_response, f"/manage/sites/{self.site.pk}/edit/")
        self.assertFalse(SiteEquipment.objects.filter(pk=equipment.pk).exists())
        self.site.refresh_from_db()
        self.assertEqual(self.site.company, "Sun Co")

    def test_non_admin_cannot_edit_sites(self):
        self.client.force_login(self.user)
        self.assertEqual(self.client.get(f"/manage/sites/{self.site.pk}/edit/").status_code, 302)

    def test_duplicate_site_day_submission_is_rejected_by_database(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        Submission.objects.create(site=self.site, template=self.template, submitted_by=admin, checklist_date=date(2026, 11, 20))
        with self.assertRaises(Exception):
            Submission.objects.create(site=self.site, template=self.template, submitted_by=admin, checklist_date=date(2026, 11, 20))

    def test_multiple_checklists_sequential_flow_for_operator(self):
        # Create a second template and assign both to self.site
        second_template = ChecklistTemplate.objects.create(name="Inverter Checks", version=1)
        item2 = ChecklistItem.objects.create(template=second_template, checkpoint_no=1, description="Grid Voltage", answer_type="number", position=1)
        SiteChecklist.objects.create(site=self.site, template=self.template, position=1)
        SiteChecklist.objects.create(site=self.site, template=second_template, position=2)

        self.client.force_login(self.user)
        # Step 1: submit the first template
        self.post_preview()
        finalize_resp = self.client.post("/checklist/finalize/", {"confirm_lockin": "yes"})
        # Should redirect to checklist-form to fill the second template
        self.assertRedirects(finalize_resp, "/checklist/")
        self.assertTrue(Submission.objects.filter(site=self.site, template=self.template).exists())
        self.assertFalse(Submission.objects.filter(site=self.site, template=second_template).exists())

        # Step 2: fill second template
        resp2 = self.client.get("/checklist/")
        self.assertEqual(resp2.status_code, 200)
        self.assertContains(resp2, "Inverter Checks")
        self.assertContains(resp2, "Grid Voltage")

        # Submit preview for second template
        post_resp = self.client.post("/checklist/", {f"item_{item2.id}": "415", "remarks": "Inverter ok"})
        self.assertEqual(post_resp.status_code, 200)
        # Lock in second template
        finalize2 = self.client.post("/checklist/finalize/", {"confirm_lockin": "yes"})
        sub2 = Submission.objects.get(site=self.site, template=second_template)
        # Should redirect to review page since all are complete
        self.assertRedirects(finalize2, f"/submissions/{sub2.pk}/review/")

    def test_zero_sequence_checklists_can_be_filled_in_any_order(self):
        second_template = ChecklistTemplate.objects.create(name="Second checklist", version=1)
        second_item = ChecklistItem.objects.create(
            template=second_template, checkpoint_no=1, description="Second check", answer_type="text", position=1
        )
        SiteChecklist.objects.create(site=self.site, template=self.template, position=0)
        SiteChecklist.objects.create(site=self.site, template=second_template, position=0)
        self.client.force_login(self.user)

        dashboard = self.client.get("/")
        self.assertEqual(dashboard.status_code, 200)
        self.assertEqual(sum(item["available"] for item in dashboard.context["checklist_status_list"]), 2)
        self.assertContains(dashboard, "Any order")
        self.assertEqual(dashboard.content.decode().count("Fill now →"), 2)

        # Select and preview the second checklist first, despite its name/order.
        entry = self.client.get(f"/checklist/?template_id={second_template.pk}")
        self.assertEqual(entry.status_code, 200)
        self.assertContains(entry, "Second check")
        preview = self.client.post("/checklist/", {"template_id": second_template.pk, f"item_{second_item.pk}": "ok"})
        self.assertEqual(preview.status_code, 200)
        lockin = self.client.get(f"/checklist/finalize/?template_id={second_template.pk}")
        self.assertEqual(lockin.status_code, 200)
        finalized = self.client.post("/checklist/finalize/", {"template_id": second_template.pk, "confirm_lockin": "yes"})
        self.assertTrue(Submission.objects.filter(site=self.site, template=second_template).exists())
        self.assertFalse(Submission.objects.filter(site=self.site, template=self.template).exists())
        self.assertRedirects(finalized, "/checklist/")

    def test_only_next_positive_sequence_is_available_and_cannot_be_skipped(self):
        second_template = ChecklistTemplate.objects.create(name="Second sequence", version=1)
        SiteChecklist.objects.create(site=self.site, template=self.template, position=1)
        SiteChecklist.objects.create(site=self.site, template=second_template, position=2)
        self.client.force_login(self.user)

        dashboard = self.client.get("/")
        statuses = {item["template"].pk: item for item in dashboard.context["checklist_status_list"]}
        self.assertTrue(statuses[self.template.pk]["available"])
        self.assertFalse(statuses[second_template.pk]["available"])
        self.assertContains(dashboard, "Fill now →")
        self.assertContains(dashboard, "Upcoming")

        blocked = self.client.get(f"/checklist/?template_id={second_template.pk}")
        self.assertRedirects(blocked, "/")
        self.assertFalse(Submission.objects.filter(site=self.site, template=second_template).exists())

    def test_unsequenced_checklists_are_available_alongside_sequence_lane(self):
        optional_template = ChecklistTemplate.objects.create(name="Any-order checklist", version=1)
        sequenced_template = ChecklistTemplate.objects.create(name="Sequenced checklist", version=1)
        later_template = ChecklistTemplate.objects.create(name="Later checklist", version=1)
        SiteChecklist.objects.create(site=self.site, template=optional_template, position=0)
        SiteChecklist.objects.create(site=self.site, template=sequenced_template, position=1)
        SiteChecklist.objects.create(site=self.site, template=later_template, position=2)
        self.client.force_login(self.user)

        dashboard = self.client.get("/")
        statuses = {item["template"].pk: item for item in dashboard.context["checklist_status_list"]}
        self.assertTrue(statuses[optional_template.pk]["available"])
        self.assertTrue(statuses[sequenced_template.pk]["available"])
        self.assertFalse(statuses[later_template.pk]["available"])
        self.assertEqual(dashboard.content.decode().count("Fill now →"), 2)

    def test_previous_ist_day_submissions_do_not_block_new_day(self):
        from unittest.mock import patch

        first_template = ChecklistTemplate.objects.create(name="Daily first", version=1)
        second_template = ChecklistTemplate.objects.create(name="Daily second", version=1)
        SiteChecklist.objects.create(site=self.site, template=first_template, position=0)
        SiteChecklist.objects.create(site=self.site, template=second_template, position=0)
        previous_day = date(2026, 11, 20)
        today = date(2026, 11, 21)
        Submission.objects.create(
            site=self.site, template=first_template, submitted_by=self.user, checklist_date=previous_day
        )
        Submission.objects.create(
            site=self.site, template=second_template, submitted_by=self.user, checklist_date=previous_day
        )
        self.client.force_login(self.user)

        # At the next IST local date, prior-day submissions no longer mark either checklist complete.
        with patch("checklist.views.timezone.localdate", return_value=today), patch(
            "checklist.views.GlobalSettings.is_open", return_value=True
        ):
            dashboard = self.client.get("/")
            self.assertEqual(dashboard.status_code, 200)
            statuses = {item["template"].pk: item for item in dashboard.context["checklist_status_list"]}
            self.assertTrue(statuses[first_template.pk]["available"])
            self.assertTrue(statuses[second_template.pk]["available"])
            self.assertIsNone(statuses[first_template.pk]["submission"])
            self.assertIsNone(statuses[second_template.pk]["submission"])

            entry = self.client.get(f"/checklist/?template_id={second_template.pk}")
            self.assertEqual(entry.status_code, 200)
            self.assertContains(entry, second_template.name)

            # The database permits a new submission for the same checklist on the new date.
            new_day_submission = Submission.objects.create(
                site=self.site, template=second_template, submitted_by=self.user, checklist_date=today
            )
            self.assertNotEqual(new_day_submission.checklist_date, previous_day)


    def test_admin_upload_excel_checklist_template(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        self.client.force_login(admin)

        wb = Workbook()
        ws = wb.active
        ws.title = "Checks"
        ws.append(["Checkpoint", "Description", "Standard", "Method", "Frequency"])
        ws.append([1, "HT Breaker Trip Test", "No trip failure", "Visual", "Daily"])
        ws.append([2, "HT Busbar Temperature", "Below 55 C", "IR Gun", "Daily"])
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        upload_file = SimpleUploadedFile("ht_panel.xlsx", buf.read(), content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        response = self.client.post("/manage/checklists/", {
            "template_name": "HT Panel Daily Inspection",
            "file": upload_file,
            "assign_to_sites": [self.site.pk],
        })
        new_template = ChecklistTemplate.objects.get(name="HT Panel Daily Inspection")
        self.assertRedirects(response, f"/manage/checklists/{new_template.pk}/preview/")
        self.assertEqual(new_template.items.count(), 2)
        item = new_template.items.get(checkpoint_no=2)
        self.assertEqual(item.answer_type, "number")
        self.assertEqual(item.unit, "°C")
        self.assertTrue(SiteChecklist.objects.filter(site=self.site, template=new_template).exists())

    def test_admin_preview_checklist_template(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        self.client.force_login(admin)

        # Admin can view preview
        resp = self.client.get(f"/manage/checklists/{self.template.pk}/preview/")
        self.assertEqual(resp.status_code, 200)
        self.assertContains(resp, self.template.name)
        self.assertContains(resp, "Winding temperature")
        self.assertContains(resp, "Yes / No")
        self.assertContains(resp, "Number")
        self.assertContains(resp, "TOTAL CHECKPOINTS")

        # Operator (non-staff) cannot view preview
        self.client.force_login(self.user)
        op_resp = self.client.get(f"/manage/checklists/{self.template.pk}/preview/")
        self.assertNotEqual(op_resp.status_code, 200)

    def test_import_uploaded_excel_multiline_headers_and_sections(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "SolarInverter"
        ws.append(["", "", "", "", "", ""])
        ws.append(["", "Section", "Check\nPoint No.", "Job Description", "Standard", "Method/\nTool", "Frequency"])
        ws.append(["", "AC Cabinet", 1, "AC Breaker Status", "Closed", "Visual", "Daily"])
        ws.append(["", "", 2, "Inverter Temperature", "Below 70 C", "Digital Display", "Daily"])
        ws.append(["", "DC Combiner", 3, "DC String Fuse Check", "Intact", "Multimeter", "Daily"])
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)

        from checklist.services.workbook_import import import_uploaded_excel
        template, count = import_uploaded_excel(buf, "Solar Inverter Station")
        self.assertEqual(count, 3)
        self.assertEqual(template.items.count(), 3)
        item1 = template.items.get(checkpoint_no=1)
        self.assertEqual(item1.section, "AC Cabinet")
        self.assertEqual(item1.description, "AC Breaker Status")
        item2 = template.items.get(checkpoint_no=2)
        self.assertEqual(item2.section, "AC Cabinet")  # inherited section
        self.assertEqual(item2.answer_type, "number")
        item3 = template.items.get(checkpoint_no=3)
        self.assertEqual(item3.section, "DC Combiner")

    def test_admin_assign_and_remove_site_checklist(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        self.client.force_login(admin)
        new_template = ChecklistTemplate.objects.create(name="Battery Bank Check", version=1)

        # Assign checklist
        assign_resp = self.client.post(f"/manage/sites/{self.site.pk}/checklists/assign/", {
            "template": str(new_template.pk),
            "position": "3",
            "active": "on",
        })
        self.assertRedirects(assign_resp, f"/manage/sites/{self.site.pk}/edit/")
        sc = SiteChecklist.objects.get(site=self.site, template=new_template)
        self.assertEqual(sc.position, 3)

        # Remove checklist
        remove_resp = self.client.post(f"/manage/sites/{self.site.pk}/checklists/{sc.pk}/remove/")
        self.assertRedirects(remove_resp, f"/manage/sites/{self.site.pk}/edit/")
        self.assertFalse(SiteChecklist.objects.filter(pk=sc.pk).exists())

    def test_admin_site_status_inline_review_tab(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        sub = Submission.objects.create(site=self.site, template=self.template, submitted_by=self.user, checklist_date=date(2026, 11, 20), remarks="Routine checks fine")
        Answer.objects.create(submission=sub, item=self.number_item, value="70.5", comment="Stable")
        self.client.force_login(admin)

        # Access inline review tab for same day or date range
        response = self.client.get(f"/?tab=review&site={self.site.pk}&start_date=2026-11-20&end_date=2026-11-20")
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "North Solar")
        self.assertContains(response, "Test Trafo")
        self.assertContains(response, "70.5")
        self.assertContains(response, "Routine checks fine")
        self.assertContains(response, "Stable")

    def test_primary_admin_deadcalm_hidden_and_protected(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        deadcalm = User.objects.get(username="DeadCalm")
        self.client.force_login(admin)

        # Ensure DeadCalm is not in manage users list
        users_resp = self.client.get("/manage/users/")
        self.assertEqual(users_resp.status_code, 200)
        self.assertNotContains(users_resp, "DeadCalm")

        # Ensure editing DeadCalm returns 404
        edit_resp = self.client.get(f"/manage/users/{deadcalm.pk}/edit/")
        self.assertEqual(edit_resp.status_code, 404)

    def test_admin_credential_change_with_valid_passcode(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        self.client.force_login(admin)

        response = self.client.post("/admin-settings/", {
            "change_credentials": "1",
            "new_username": "site_admin",
            "new_password": "new-strong-password",
            "confirm_password": "new-strong-password",
            "verification_passcode": "TalfSolar",
        })
        self.assertRedirects(response, "/admin-settings/")
        admin.refresh_from_db()
        self.assertEqual(admin.username, "site_admin")
        self.assertTrue(admin.check_password("new-strong-password"))

    def test_admin_credential_change_fails_with_invalid_passcode(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        self.client.force_login(admin)

        response = self.client.post("/admin-settings/", {
            "change_credentials": "1",
            "new_username": "hacked_admin",
            "new_password": "hacked-password",
            "confirm_password": "hacked-password",
            "verification_passcode": "wrong_passcode",
        })
        self.assertEqual(response.status_code, 200)
        admin.refresh_from_db()
        self.assertEqual(admin.username, "admin")
        self.assertFalse(admin.check_password("hacked-password"))

    def test_admin_cannot_rename_to_deadcalm(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        self.client.force_login(admin)

        response = self.client.post("/admin-settings/", {
            "change_credentials": "1",
            "new_username": "DeadCalm",
            "new_password": "",
            "confirm_password": "",
            "verification_passcode": "TalfSolar",
        })
        self.assertEqual(response.status_code, 200)
        admin.refresh_from_db()
        self.assertEqual(admin.username, "admin")


