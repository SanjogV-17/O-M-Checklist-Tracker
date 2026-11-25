import io
import re
from pathlib import Path

from django.db import transaction
from openpyxl import load_workbook

from checklist.models import ChecklistItem, ChecklistTemplate, Site, SiteEquipment


def _cell_text(value):
    return "" if value is None else str(value).strip()


def _answer_type(description, standard):
    text = f"{description} {standard}".lower()
    if "temperature" in text:
        return ChecklistItem.AnswerType.NUMBER
    if any(word in text for word in ("should be", "no leakage", "no sign", "should open", "should be clean", "no dirt", "no any sign")):
        return ChecklistItem.AnswerType.BOOLEAN
    return ChecklistItem.AnswerType.TEXT


def import_workbook(path, site_name=None):
    """Import checklist definitions from the sample workbook, ignoring daily values."""
    path = Path(path)
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook["spv"] if "spv" in workbook.sheetnames else workbook.active
    document_number = ""
    company = ""
    equipment = ""
    for row in sheet.iter_rows(min_row=1, max_row=5, values_only=True):
        for value in row:
            text = _cell_text(value)
            if text.startswith("Document No."):
                document_number = text.partition(":")[2].strip()
            elif "Energy India" in text or "Pvt Ltd" in text:
                company = text
            elif text.startswith("Equipment Name:"):
                equipment = text.partition(":")[2].strip()

    template, _ = ChecklistTemplate.objects.get_or_create(
        name="Trafo Daily Checklist", version=1,
        defaults={"document_number": document_number, "company": company, "equipment_name": equipment},
    )
    if template.items.exists():
        return template, 0

    current_section = ""
    created = 0
    for row in sheet.iter_rows(min_row=7, max_col=7, values_only=True):
        section, checkpoint_no, description, standard, method, frequency = row[1], row[2], row[3], row[4], row[5], row[6]
        if section:
            current_section = _cell_text(section)
        if not checkpoint_no or not description:
            continue
        description = _cell_text(description)
        standard = _cell_text(standard)
        kind = _answer_type(description, standard)
        unit = "°C" if "temperature" in description.lower() else ""
        ChecklistItem.objects.create(
            template=template,
            section=current_section,
            checkpoint_no=int(checkpoint_no),
            description=description,
            standard=standard,
            method=_cell_text(method),
            frequency=_cell_text(frequency),
            answer_type=kind,
            unit=unit,
            position=int(checkpoint_no),
        )
        created += 1

    if site_name:
        site, _ = Site.objects.get_or_create(
            name=site_name,
            defaults={"company": company, "location": "Jhatikara DL9"},
        )
        if equipment:
            SiteEquipment.objects.get_or_create(site=site, name=equipment)
    workbook.close()
    return template, created


def _norm(v):
    return re.sub(r'[\s\n_/]+', ' ', _cell_text(v)).lower().strip()


# Substring hint lists for header detection
_CP_HINTS = ["checkpoint", "check point", "point no", "s.no", "sl.no", "sr.no",
             "sl no", "sr no", "s/n", "item no", "no.", "no", "sl", "sr", "#"]
_DESC_HINTS = ["description", "job description", "observation", "activity",
               "parameter", "particulars", "check", "item", "title"]
_STD_HINTS = ["standard", "specification", "acceptance", "criteria", "limits", "norm", "requirement"]
_MTH_HINTS = ["method", "tool", "instrument", "procedure"]
_FRQ_HINTS = ["frequency", "freq", "period", "interval"]
_SEC_HINTS = ["section", "category", "system", "subsystem", "area", "group", "equipment"]
_STATUS_HINTS = ["operational", "status", "condition", "working", "inspection", "check", "state", "health", "result", "ok"]

# Strings that mark a header-only row (not a data row)
_HEADER_KEYWORDS = frozenset([
    "description", "job description", "observation", "check", "item",
    "checkpoint", "check point", "standard", "specification", "method", "tool",
    "frequency", "freq", "section", "category", "s.no", "sl.no", "sr.no", "sl", "sr", "no",
])


def _get_dropdown_cols(ws):
    """
    Detect Excel columns having list / dropdown data validation.
    Returns a dict mapping 0-indexed column index → option string
    (comma-separated values extracted from the validation formula).
    Also handles:
      - Inline list:  "\"Yes,No,OK\""  → "Yes,No,OK"
      - Named range / cell ref: returns the raw formula1 string as-is
    """
    col_options = {}  # col_idx → options string
    if ws and hasattr(ws, "data_validations"):
        for dv in ws.data_validations.dataValidation:
            if dv.type == "list":
                f1 = (dv.formula1 or "").strip()
                # Strip surrounding quotes for inline lists: "\"A,B,C\"" → "A,B,C"
                if f1.startswith('"') and f1.endswith('"'):
                    options_str = f1[1:-1]
                else:
                    options_str = f1  # named range or cell ref — store raw
                for rng in dv.sqref.ranges:
                    for col_idx in range(rng.min_col, rng.max_col + 1):
                        col_options[col_idx - 1] = options_str  # 0-indexed
    return col_options


def _classify_options(options_str):
    """
    Given a comma-separated option string, return (answer_type, choices_str):
    - If only Yes/No/OK/Pass/True/False variants → BOOLEAN, choices_str=""
    - Otherwise → CHOICES, choices_str=options_str
    """
    if not options_str:
        return ChecklistItem.AnswerType.BOOLEAN, ""
    opts_lower = {o.strip().lower() for o in options_str.split(",")}
    boolean_words = {"yes", "no", "ok", "pass", "fail", "true", "false", "y", "n"}
    if opts_lower and opts_lower.issubset(boolean_words):
        return ChecklistItem.AnswerType.BOOLEAN, ""
    return ChecklistItem.AnswerType.CHOICES, options_str


def _is_header_row(row_texts):
    """Return True if the row looks like a header (no useful data)."""
    non_empty = [t for t in row_texts if t]
    if not non_empty:
        return False
    has_cp = any(any(h in t for h in _CP_HINTS) for t in non_empty)
    has_desc = any(any(h in t for h in _DESC_HINTS) for t in non_empty)
    keyword_hits = sum(1 for t in non_empty if t in _HEADER_KEYWORDS or any(h in t for h in _DESC_HINTS + _CP_HINTS))
    numeric_cells = sum(1 for t in non_empty if t.isdigit() or (t.replace('.', '', 1).isdigit() and t.count('.') <= 1))
    return has_cp and has_desc and keyword_hits >= 2 and numeric_cells == 0


def import_uploaded_excel(file_obj, template_name):
    """
    Parse an uploaded Excel file into a new ChecklistTemplate.
    Supports both:
    1) Standard inspection checklists with Checkpoint / Description / Standard columns (e.g. Trafo).
    2) Equipment / component status tables with dropdown or operational checks (e.g. Dhansa Inverter1).
    Returns (template, items_created).
    """
    if hasattr(file_obj, "read"):
        file_bytes = file_obj.read()
        if hasattr(file_obj, "seek"):
            file_obj.seek(0)
    else:
        with open(file_obj, "rb") as fp:
            file_bytes = fp.read()

    wb_data = load_workbook(io.BytesIO(file_bytes), data_only=True)
    try:
        wb_val = load_workbook(io.BytesIO(file_bytes), data_only=False)
    except Exception:
        wb_val = None

    # Determine candidate sheets order: prefer 'spv' if present, then all other sheets
    sheets_to_process = []
    if "spv" in wb_data.sheetnames:
        sheets_to_process.append("spv")
    for name in wb_data.sheetnames:
        if name != "spv":
            sheets_to_process.append(name)

    # First check which sheets match Strategy 1 (Standard Inspection Checklist)
    strategy1_sheets = set()
    for name in sheets_to_process:
        s = wb_data[name]
        for ri, row in enumerate(s.iter_rows(min_row=1, max_row=15, values_only=True), start=1):
            texts = [_norm(v) for v in row]
            if any(any(h in t for h in _CP_HINTS) for t in texts if t) and any(any(h in t for h in _DESC_HINTS) for t in texts if t):
                strategy1_sheets.add(name)
                break

    parsed_items = []
    checkpoint_counter = 0
    seen_checkpoints = set()

    for sheet_name in sheets_to_process:
        sheet = wb_data[sheet_name]
        ws_val = wb_val[sheet_name] if wb_val and sheet_name in wb_val.sheetnames else None
        dropdown_cols = _get_dropdown_cols(ws_val)

        if sheet_name in strategy1_sheets:
            # STRATEGY 1: Standard Checklist (Trafo style)
            col_map = {}
            header_row = None
            for ri, row in enumerate(sheet.iter_rows(min_row=1, max_row=15, values_only=True), start=1):
                texts = [_norm(v) for v in row]
                has_cp = any(any(h in t for h in _CP_HINTS) for t in texts if t)
                has_desc = any(any(h in t for h in _DESC_HINTS) for t in texts if t)
                if has_cp and has_desc:
                    header_row = ri
                    for ci, t in enumerate(texts):
                        if "checkpoint" not in col_map and any(h in t for h in _CP_HINTS):
                            col_map["checkpoint"] = ci
                        elif "description" not in col_map and any(h in t for h in _DESC_HINTS):
                            col_map["description"] = ci
                        elif "standard" not in col_map and any(h in t for h in _STD_HINTS):
                            col_map["standard"] = ci
                        elif "method" not in col_map and any(h in t for h in _MTH_HINTS):
                            col_map["method"] = ci
                        elif "frequency" not in col_map and any(h in t for h in _FRQ_HINTS):
                            col_map["frequency"] = ci
                        elif "section" not in col_map and any(h in t for h in _SEC_HINTS):
                            col_map["section"] = ci
                    break

            current_section = ""
            for row in sheet.iter_rows(min_row=(header_row or 1) + 1, values_only=True):
                row = list(row)
                if not any(v is not None and _cell_text(v) for v in row):
                    continue
                row_texts = [_norm(v) for v in row]
                if _is_header_row(row_texts):
                    continue

                def gcol(k):
                    idx = col_map.get(k)
                    return _cell_text(row[idx]) if idx is not None and idx < len(row) else ""

                sec = gcol("section")
                if sec and not sec.isdigit():
                    current_section = sec

                desc = gcol("description")
                if not desc or _norm(desc) in _HEADER_KEYWORDS:
                    continue

                cp_raw = gcol("checkpoint")
                try:
                    cp = int(float(cp_raw)) if cp_raw else None
                except (ValueError, TypeError):
                    cp = None

                if cp is None or cp <= 0 or cp in seen_checkpoints:
                    checkpoint_counter += 1
                    cp = checkpoint_counter
                else:
                    checkpoint_counter = max(checkpoint_counter, cp)
                seen_checkpoints.add(cp)

                std = gcol("standard")
                mth = gcol("method")
                frq = gcol("frequency")

                # Check if this column has a dropdown validation
                desc_col_idx = col_map.get("description")
                dropdown_opts = dropdown_cols.get(desc_col_idx, "") if desc_col_idx is not None else ""
                if dropdown_opts:
                    kind, choices_str = _classify_options(dropdown_opts)
                else:
                    kind = _answer_type(desc, std)
                    choices_str = ""

                unit = "°C" if "temperature" in desc.lower() else ""

                parsed_items.append({
                    "checkpoint_no": cp,
                    "section": current_section or sheet_name,
                    "description": desc,
                    "standard": std,
                    "method": mth,
                    "frequency": frq,
                    "answer_type": kind,
                    "choices": choices_str,
                    "unit": unit,
                    "position": cp,
                })

        else:
            # STRATEGY 2: Equipment / Component Matrix Table (e.g. Dhansa Inverter1)
            matrix_header_row = None
            headers = []
            for ri, row in enumerate(sheet.iter_rows(min_row=1, max_row=5, values_only=True), start=1):
                row_texts = [_cell_text(v) for v in row]
                non_empty = [t for t in row_texts if t]
                if len(non_empty) >= 2:
                    matrix_header_row = ri
                    headers = row_texts
                    break

            if matrix_header_row is None:
                continue

            # Identify check / status column — now returns dict of col_idx → options_str
            check_col = None
            check_col_options = ""
            for ci in range(len(headers)):
                if ci in dropdown_cols:
                    check_col = ci
                    check_col_options = dropdown_cols[ci]
                    break
            if check_col is None:
                for ci, h in enumerate(headers):
                    if any(sh in _norm(h) for sh in _STATUS_HINTS):
                        check_col = ci
                        break
            if check_col is None and len(headers) > 1:
                for ci in range(len(headers) - 1, -1, -1):
                    if headers[ci]:
                        check_col = ci
                        break

            col_0 = 0
            while col_0 < len(headers) and not headers[col_0]:
                col_0 += 1

            col_0_header = headers[col_0] if col_0 < len(headers) else ""
            if col_0_header and _norm(col_0_header) not in ["no", "sl", "s.no", "item", "#"]:
                section_name = col_0_header
            else:
                section_name = sheet_name

            check_name = headers[check_col] if check_col is not None and check_col < len(headers) and headers[check_col] else "Operational"
            has_dropdown = (check_col in dropdown_cols) or any(sh in _norm(check_name) for sh in _STATUS_HINTS)

            # Classify the dropdown options once for the whole column
            if check_col_options:
                col_kind, col_choices = _classify_options(check_col_options)
            elif has_dropdown:
                # Status-hint column without explicit validation → default Yes/No
                col_kind, col_choices = ChecklistItem.AnswerType.BOOLEAN, ""
            else:
                col_kind, col_choices = None, ""  # will fall through to _answer_type

            for row in sheet.iter_rows(min_row=matrix_header_row + 1, values_only=True):
                row = list(row)
                if not any(v is not None and _cell_text(v) for v in row):
                    continue

                item_val = _cell_text(row[col_0]) if col_0 < len(row) else ""
                if not item_val or _norm(item_val) in [_norm(h) for h in headers if h]:
                    continue
                if _norm(item_val) in ["total", "grand total", "summary", "subtotal", "average", "avg"]:
                    continue

                checkpoint_counter += 1
                seen_checkpoints.add(checkpoint_counter)

                # Assemble specifications from intermediate columns
                specs = []
                for ci in range(col_0 + 1, len(row)):
                    if ci == check_col:
                        continue
                    if ci < len(headers) and headers[ci] and _cell_text(row[ci]):
                        specs.append(f"{headers[ci]}: {_cell_text(row[ci])}")
                standard_str = " | ".join(specs)

                if check_name.lower() in item_val.lower():
                    desc = item_val
                else:
                    desc = f"{item_val} - {check_name} Check"

                if col_kind is not None:
                    kind = col_kind
                    choices_str = col_choices
                else:
                    kind = _answer_type(desc, standard_str)
                    choices_str = ""

                parsed_items.append({
                    "checkpoint_no": checkpoint_counter,
                    "section": section_name,
                    "description": desc,
                    "standard": standard_str,
                    "method": "Visual / SCADA",
                    "frequency": "Daily",
                    "answer_type": kind,
                    "choices": choices_str,
                    "unit": "",
                    "position": checkpoint_counter,
                })

    wb_data.close()
    if wb_val:
        wb_val.close()

    # Create template and items atomically
    with transaction.atomic():
        existing = ChecklistTemplate.objects.filter(name=template_name).order_by("-version").first()
        version = (existing.version + 1) if existing else 1
        template = ChecklistTemplate.objects.create(name=template_name, version=version)

        items_to_create = [
            ChecklistItem(
                template=template,
                section=it["section"],
                checkpoint_no=it["checkpoint_no"],
                description=it["description"],
                standard=it["standard"],
                method=it["method"],
                frequency=it["frequency"],
                answer_type=it["answer_type"],
                choices=it.get("choices", ""),
                unit=it["unit"],
                position=it["position"],
            )
            for it in parsed_items
        ]
        ChecklistItem.objects.bulk_create(items_to_create)

    return template, len(items_to_create)


def import_uploaded_pdf(file_obj, template_name):
    """
    Parse an uploaded PDF into a ChecklistTemplate using pdfplumber if available,
    falling back to a single placeholder item with instructions.
    Returns (template, items_created).
    """
    existing = ChecklistTemplate.objects.filter(name=template_name).order_by("-version").first()
    version = (existing.version + 1) if existing else 1
    template = ChecklistTemplate.objects.create(name=template_name, version=version)

    try:
        import pdfplumber  # optional dependency
    except ImportError:
        ChecklistItem.objects.create(
            template=template,
            checkpoint_no=1,
            description="[PDF import not available — install pdfplumber or re-upload as Excel]",
            answer_type=ChecklistItem.AnswerType.TEXT,
            position=1,
        )
        return template, 1

    created = 0
    checkpoint_counter = 0
    current_section = ""

    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row:
                        continue
                    row = [_cell_text(c) for c in row]
                    desc = row[1] if len(row) > 1 else row[0]
                    if not desc or desc.lower() in ("description", "observation", "job description"):
                        continue
                    checkpoint_counter += 1
                    standard = row[2] if len(row) > 2 else ""
                    kind = _answer_type(desc, standard)
                    ChecklistItem.objects.create(
                        template=template,
                        section=current_section,
                        checkpoint_no=checkpoint_counter,
                        description=desc,
                        standard=standard,
                        answer_type=kind,
                        position=checkpoint_counter,
                    )
                    created += 1

    if created == 0:
        ChecklistItem.objects.create(
            template=template,
            checkpoint_no=1,
            description="[No table rows found in PDF — please re-upload as Excel for best results]",
            answer_type=ChecklistItem.AnswerType.TEXT,
            position=1,
        )
        created = 1

    return template, created
