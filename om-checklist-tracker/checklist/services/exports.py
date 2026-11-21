from io import BytesIO

from django.http import HttpResponse
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from checklist.models import Answer


def _equipment_label(submission):
    return submission.equipment_snapshot or ", ".join(submission.site.equipment.values_list("name", flat=True)) or submission.template.equipment_name


def excel_response(submission):
    template_path = __import__("pathlib").Path(__file__).resolve().parents[2] / "Trafo_Checklist_Nov_2026_Populated.xlsx"
    use_trafo_sheet = template_path.exists() and "trafo" in submission.template.name.lower()
    workbook = load_workbook(template_path) if use_trafo_sheet else None
    if workbook is None:
        from openpyxl import Workbook
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Checklist"
        sheet.append(["Section", "Checkpoint", "Job Description", "Standard", "Method/Tool", "Frequency", "Value", "Comment"])
        items = list(submission.template.items.all())
        answer_by_item = {a.item_id: a for a in submission.answers.select_related("item")}
        for item in items:
            ans = answer_by_item.get(item.id)
            val = ans.value if ans else ""
            com = ans.comment if ans else ""
            sheet.append([item.section, item.checkpoint_no, item.description, item.standard, item.method, item.frequency, val, com])
    else:
        sheet = workbook["spv"] if "spv" in workbook.sheetnames else workbook.active
        # Remove populated sample daily values; exports contain only live submitted data.
        for row in range(7, sheet.max_row + 1):
            for column in range(8, sheet.max_column + 1):
                cell = sheet.cell(row, column)
                if cell.__class__.__name__ != "MergedCell":
                    cell.value = None
        day_column = 8 + submission.checklist_date.day - 1
        sheet["C3"] = f"Document No. {submission.template.document_number}"
        sheet["E3"] = submission.site.company or submission.template.company
        sheet["H4"] = f"Equipment Name: {_equipment_label(submission)}"
        sheet["H5"] = f"Month - {submission.checklist_date:%b-%Y}"
        for day in range(1, 32):
            sheet.cell(6, 7 + day, day)
        answer_by_item = {a.item_id: a for a in submission.answers.select_related("item")}
        items = list(submission.template.items.all())
        for item in items:
            row = item.position + 6
            answer = answer_by_item.get(item.id)
            if answer and row <= sheet.max_row:
                cell = sheet.cell(row, day_column, answer.value)
                cell.comment = None
        sheet["B35"] = f"Remarks: {submission.remarks}"
        sheet["C3"].comment = None
        sheet["B36"] = f"Submitted by: {submission.signer_name or submission.submitted_by.get_full_name() or submission.submitted_by.username}"
        sheet["E36"] = f"Submitted at: {submission.submitted_at.strftime('%Y-%m-%d %H:%M %Z')}"

    # Keep comments and remarks visible without relying on space in the imported workbook layout.
    review = workbook.create_sheet("Submission review")
    review.append(["Checkpoint", "Observation", "Reading / choice", "Comment"])
    for answer in submission.answers.select_related("item").order_by("item__position", "item__checkpoint_no"):
        review.append([answer.item.checkpoint_no, answer.item.description, answer.value, answer.comment])
    review.append([])
    review.append(["Overall remarks", submission.remarks])
    review.append(["Submitted by", submission.signer_name or submission.submitted_by.get_full_name() or submission.submitted_by.username])
    review.append(["Submitted at", submission.submitted_at.strftime("%Y-%m-%d %H:%M %Z")])

    output = BytesIO()
    workbook.save(output)
    response = HttpResponse(output.getvalue(), content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response["Content-Disposition"] = f'attachment; filename="{submission.site.name.replace(" ", "_")}_{submission.checklist_date}.xlsx"'
    return response


def pdf_response(submission):
    output = BytesIO()
    doc = SimpleDocTemplate(output, pagesize=landscape(A4), rightMargin=28, leftMargin=28, topMargin=28, bottomMargin=28)
    styles = getSampleStyleSheet()
    story = [Paragraph("O and M Checklist Tracker", styles["Title"]), Spacer(1, 8)]
    story.append(Paragraph(f"{submission.site.name} — {submission.checklist_date:%d %B %Y}", styles["Heading2"]))
    story.append(Paragraph(f"Equipment: {_equipment_label(submission)} | Submitted by: {submission.signer_name or submission.submitted_by.username} | Submitted: {submission.submitted_at:%H:%M IST}", styles["Normal"]))
    story.append(Spacer(1, 12))
    data = [["#", "Section", "Observation", "Standard", "Reading / Status", "Comment"]]
    for answer in submission.answers.select_related("item").order_by("item__position"):
        item = answer.item
        data.append([str(item.checkpoint_no), item.section, Paragraph(item.description, styles["BodyText"]), Paragraph(item.standard or "—", styles["BodyText"]), answer.value or "—", Paragraph(answer.comment or "—", styles["BodyText"])])
    table = Table(data, colWidths=[28, 90, 220, 170, 100, 150], repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#123b32")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cbd5d1")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f1f5f3")]),
        ("PADDING", (0, 0), (-1, -1), 5),
    ]))
    story.extend([table, Spacer(1, 10), Paragraph(f"Remarks: {submission.remarks or '—'}", styles["Normal"])])
    doc.build(story)
    response = HttpResponse(output.getvalue(), content_type="application/pdf")
    response["Content-Disposition"] = f'attachment; filename="{submission.site.name.replace(" ", "_")}_{submission.checklist_date}.pdf"'
    return response
