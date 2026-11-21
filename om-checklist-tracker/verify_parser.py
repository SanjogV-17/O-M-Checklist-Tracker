import io
import re
from openpyxl import load_workbook

def _cell_text(v):
    return "" if v is None else str(v).strip()

def _norm(v):
    return re.sub(r'[\s\n_/]+', ' ', _cell_text(v)).lower().strip()

_CP_HINTS = ["checkpoint", "check point", "point no", "s.no", "sl.no", "sr.no",
             "sl no", "sr no", "s/n", "item no", "no.", "no", "sl", "sr", "#"]
_DESC_HINTS = ["description", "job description", "observation", "activity",
               "parameter", "particulars", "check", "item", "title"]
_STD_HINTS = ["standard", "specification", "acceptance", "criteria", "limits", "norm", "requirement"]
_MTH_HINTS = ["method", "tool", "instrument", "procedure"]
_FRQ_HINTS = ["frequency", "freq", "period", "interval"]
_SEC_HINTS = ["section", "category", "system", "subsystem", "area", "group", "equipment"]
_STATUS_HINTS = ["operational", "status", "condition", "working", "inspection", "check", "state", "health", "result", "ok"]
_HEADER_KEYWORDS = frozenset([
    "description", "job description", "observation", "check", "item",
    "checkpoint", "check point", "standard", "specification", "method", "tool",
    "frequency", "freq", "section", "category", "s.no", "sl.no", "sr.no", "sl", "sr", "no",
])

def _get_dropdown_cols(ws):
    cols = set()
    if ws and hasattr(ws, "data_validations"):
        for dv in ws.data_validations.dataValidation:
            if dv.type == "list":
                f1 = (dv.formula1 or "").lower()
                if any(w in f1 for w in ["yes", "no", "ok", "pass", "true", "false"]):
                    for rng in dv.sqref.ranges:
                        for col_idx in range(rng.min_col, rng.max_col + 1):
                            cols.add(col_idx - 1)  # 0-indexed
    return cols

def _is_header_row(row_texts):
    non_empty = [t for t in row_texts if t]
    if not non_empty:
        return False
    has_cp = any(any(h in t for h in _CP_HINTS) for t in non_empty)
    has_desc = any(any(h in t for h in _DESC_HINTS) for t in non_empty)
    keyword_hits = sum(1 for t in non_empty if t in _HEADER_KEYWORDS or any(h in t for h in _DESC_HINTS + _CP_HINTS))
    numeric_cells = sum(1 for t in non_empty if t.isdigit() or (t.replace('.', '', 1).isdigit() and t.count('.') <= 1))
    return has_cp and has_desc and keyword_hits >= 2 and numeric_cells == 0

def _answer_type(desc, standard):
    text = f"{desc} {standard}".lower()
    if any(w in text for w in ["temperature", "voltage", "current", "pressure", "kw", "hz", "°c"]):
        return "number"
    if any(w in text for w in ["should be", "no leakage", "no sign", "should open", "should be clean", "no dirt", "normal", "intact", "closed", "operational"]):
        return "boolean"
    return "text"

def parse_excel_to_items(file_bytes):
    wb_data = load_workbook(io.BytesIO(file_bytes), data_only=True)
    try:
        wb_val = load_workbook(io.BytesIO(file_bytes), data_only=False)
    except Exception:
        wb_val = None

    items = []
    checkpoint_counter = 0
    seen_checkpoints = set()

    # Determine candidate sheets
    sheets_to_process = []
    if "spv" in wb_data.sheetnames:
        sheets_to_process.append("spv")
    for name in wb_data.sheetnames:
        if name != "spv":
            sheets_to_process.append(name)

    # First pass: check if any sheet has Strategy 1 (Standard Checklist) header
    strategy1_sheets = set()
    for name in sheets_to_process:
        sheet = wb_data[name]
        for ri, row in enumerate(sheet.iter_rows(min_row=1, max_row=15, values_only=True), start=1):
            texts = [_norm(v) for v in row]
            if any(any(h in t for h in _CP_HINTS) for t in texts if t) and any(any(h in t for h in _DESC_HINTS) for t in texts if t):
                strategy1_sheets.add(name)
                break

    for sheet_name in sheets_to_process:
        sheet = wb_data[sheet_name]
        ws_val = wb_val[sheet_name] if wb_val and sheet_name in wb_val.sheetnames else None
        dropdown_cols = _get_dropdown_cols(ws_val)

        if sheet_name in strategy1_sheets:
            # STRATEGY 1: Standard Checklist Table (Trafo style)
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
                kind = _answer_type(desc, std)
                unit = "°C" if "temperature" in desc.lower() else ""

                items.append({
                    "checkpoint_no": cp,
                    "section": current_section or sheet_name,
                    "description": desc,
                    "standard": std,
                    "method": mth,
                    "frequency": frq,
                    "answer_type": kind,
                    "unit": unit,
                    "position": cp,
                })

        else:
            # STRATEGY 2: Matrix / Equipment / Dropdown Table (Dhansa Inverter1 style)
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

            # Identify check/status column
            check_col = None
            for ci in range(len(headers)):
                if ci in dropdown_cols:
                    check_col = ci
                    break
            if check_col is None:
                for ci, h in enumerate(headers):
                    if any(sh in _norm(h) for sh in _STATUS_HINTS):
                        check_col = ci
                        break
            if check_col is None and len(headers) > 1:
                # Find last column with non-empty header
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
            is_dropdown = (check_col in dropdown_cols) or any(sh in _norm(check_name) for sh in _STATUS_HINTS)

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

                kind = "boolean" if is_dropdown else _answer_type(desc, standard_str)

                items.append({
                    "checkpoint_no": checkpoint_counter,
                    "section": section_name,
                    "description": desc,
                    "standard": standard_str,
                    "method": "Visual / SCADA",
                    "frequency": "Daily",
                    "answer_type": kind,
                    "unit": "",
                    "position": checkpoint_counter,
                })

    return items

if __name__ == "__main__":
    import os
    path1 = "Dhansa Inverter1.xlsx" if os.path.exists("Dhansa Inverter1.xlsx") else "../Dhansa Inverter1.xlsx"
    with open(path1, "rb") as f:
        inv_items = parse_excel_to_items(f.read())
    print(f"Dhansa Inverter1: {len(inv_items)} items parsed.")
    assert len(inv_items) == 24
    for it in inv_items[:3]:
        print("  ", it["checkpoint_no"], it["section"], it["description"], "|", it["standard"], "|", it["answer_type"])
    print("  ...")
    for it in inv_items[-2:]:
        print("  ", it["checkpoint_no"], it["section"], it["description"], "|", it["standard"], "|", it["answer_type"])

    path2 = "Trafo_Checklist_Nov_2026_Populated.xlsx" if os.path.exists("Trafo_Checklist_Nov_2026_Populated.xlsx") else "om-checklist-tracker/Trafo_Checklist_Nov_2026_Populated.xlsx"
    with open(path2, "rb") as f:
        trafo_items = parse_excel_to_items(f.read())
    print(f"Trafo: {len(trafo_items)} items parsed.")
    assert len(trafo_items) == 28
    for it in trafo_items[:2]:
        print("  ", it["checkpoint_no"], it["section"], it["description"][:35], "|", it["standard"][:25], "|", it["answer_type"])
    print("  ...")
    for it in trafo_items[-2:]:
        print("  ", it["checkpoint_no"], it["section"], it["description"][:35], "|", it["standard"][:25], "|", it["answer_type"])

    print("ALL VERIFICATIONS PASSED!")
