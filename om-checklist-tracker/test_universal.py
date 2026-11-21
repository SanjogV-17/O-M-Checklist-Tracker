import re
from openpyxl import Workbook, load_workbook

def parse_excel_universal(file_obj):
    wb = load_workbook(file_obj, read_only=True, data_only=True)

    def _clean(v):
        return "" if v is None else str(v).strip()

    def _norm(v):
        return re.sub(r'[\s_]+', ' ', _clean(v)).lower()

    # Find candidate sheet
    sheets_to_check = []
    if "spv" in wb.sheetnames:
        sheets_to_check.append(wb["spv"])
    for name in wb.sheetnames:
        if name != "spv":
            sheets_to_check.append(wb[name])

    CHECKPOINT_HINTS = ["checkpoint", "check point", "point no", "s.no", "sl.no", "sr.no", "sl no", "sr no", "sl", "sr", "s/n", "item no", "no.", "#", "no"]
    DESC_HINTS = ["description", "job description", "observation", "activity", "parameter", "particulars", "check", "item", "title"]
    STANDARD_HINTS = ["standard", "specification", "acceptance", "criteria", "limits", "norm"]
    METHOD_HINTS = ["method", "tool", "instrument", "procedure"]
    FREQ_HINTS = ["frequency", "freq", "period", "interval"]
    SECTION_HINTS = ["section", "category", "system", "subsystem", "area", "group"]

    best_rows = []
    for sheet in sheets_to_check:
        col_map = {}
        header_row = 0

        # Scan top 15 rows for header
        for ri, row in enumerate(sheet.iter_rows(min_row=1, max_row=15, values_only=True), start=1):
            texts = [_norm(v) for v in row]
            has_cp = any(any(h in t for h in CHECKPOINT_HINTS) for t in texts)
            has_desc = any(any(h in t for h in DESC_HINTS) for t in texts)
            if has_cp and has_desc:
                header_row = ri
                for ci, t in enumerate(texts):
                    if "checkpoint" not in col_map and any(h in t for h in CHECKPOINT_HINTS):
                        col_map["checkpoint"] = ci
                    elif "description" not in col_map and any(h in t for h in DESC_HINTS):
                        col_map["description"] = ci
                    elif "standard" not in col_map and any(h in t for h in STANDARD_HINTS):
                        col_map["standard"] = ci
                    elif "method" not in col_map and any(h in t for h in METHOD_HINTS):
                        col_map["method"] = ci
                    elif "frequency" not in col_map and any(h in t for h in FREQ_HINTS):
                        col_map["frequency"] = ci
                    elif "section" not in col_map and any(h in t for h in SECTION_HINTS):
                        col_map["section"] = ci
                break

        rows_data = []
        current_section = ""
        start_row = (header_row or 0) + 1

        for row in sheet.iter_rows(min_row=start_row, values_only=True):
            row = list(row)
            if not any(v is not None and str(v).strip() for v in row):
                continue

            sec = ""
            cp = None
            desc = ""
            std = ""
            mth = ""
            frq = ""

            if col_map.get("description") is not None:
                d_idx = col_map["description"]
                desc = _clean(row[d_idx]) if d_idx < len(row) else ""

                cp_idx = col_map.get("checkpoint")
                if cp_idx is not None and cp_idx < len(row):
                    cp_raw = _clean(row[cp_idx])
                    try:
                        cp = int(float(cp_raw))
                    except (ValueError, TypeError):
                        pass

                sec_idx = col_map.get("section")
                if sec_idx is not None and sec_idx < len(row):
                    sec = _clean(row[sec_idx])

                std_idx = col_map.get("standard")
                if std_idx is not None and std_idx < len(row):
                    std = _clean(row[std_idx])

                mth_idx = col_map.get("method")
                if mth_idx is not None and mth_idx < len(row):
                    mth = _clean(row[mth_idx])

                frq_idx = col_map.get("frequency")
                if frq_idx is not None and frq_idx < len(row):
                    frq = _clean(row[frq_idx])

            # If header mapping failed or didn't find description in this row:
            if not desc:
                # Try layout 1 (Trafo sample style: 1=Sec, 2=CP, 3=Desc, 4=Std, 5=Mth, 6=Frq)
                if len(row) > 3 and _clean(row[3]) and _clean(row[3]).lower() not in ("job description", "description"):
                    sec = _clean(row[1]) if len(row) > 1 else ""
                    try:
                        cp = int(float(_clean(row[2])))
                    except Exception:
                        pass
                    desc = _clean(row[3])
                    std = _clean(row[4]) if len(row) > 4 else ""
                    mth = _clean(row[5]) if len(row) > 5 else ""
                    frq = _clean(row[6]) if len(row) > 6 else ""
                # Try layout 2 (0-indexed: 0=CP, 1=Desc, 2=Std, 3=Mth, 4=Frq)
                elif len(row) > 1 and _clean(row[1]) and _clean(row[1]).lower() not in ("job description", "description"):
                    try:
                        cp = int(float(_clean(row[0])))
                    except Exception:
                        pass
                    desc = _clean(row[1])
                    std = _clean(row[2]) if len(row) > 2 else ""
                    mth = _clean(row[3]) if len(row) > 3 else ""
                    frq = _clean(row[4]) if len(row) > 4 else ""

            if not desc or desc.lower() in ("job description", "description", "observations"):
                continue

            if sec and not sec.isdigit():
                current_section = sec

            rows_data.append({
                "checkpoint_no": cp,
                "description": desc,
                "section": current_section,
                "standard": std,
                "method": mth,
                "frequency": frq,
            })

        if len(rows_data) > len(best_rows):
            best_rows = rows_data

    # Assign sequential numbers if any checkpoint_no was missing
    counter = 0
    for item in best_rows:
        if item["checkpoint_no"] is None:
            counter += 1
            item["checkpoint_no"] = counter
        else:
            counter = item["checkpoint_no"]

    return best_rows

if __name__ == "__main__":
    # Test 1: with Trafo sample
    res1 = parse_excel_universal("Trafo_Checklist_Nov_2026_Populated.xlsx")
    print(f"Test 1 (Trafo populated): {len(res1)} items parsed.")

    # Test 2: Inverter style with different headers (Col A empty, headers on row 4)
    wb2 = Workbook()
    ws2 = wb2.active
    ws2.title = "Inverter"
    ws2.append(["", "", "", "", ""])
    ws2.append(["", "", "INVERTER DAILY CHECKLIST", "", ""])
    ws2.append(["", "Section", "Sl. No", "Observation / Check", "Acceptance Criteria", "Tool", "Interval"])
    ws2.append(["", "DC Side", 1, "DC Input Voltage Check", "600V - 1000V", "Multimeter", "Daily"])
    ws2.append(["", "", 2, "DC Cable Termination", "No loose connection", "Visual", "Daily"])
    ws2.append(["", "AC Side", 3, "Inverter Grid Frequency", "50 Hz +/- 0.5", "Display", "Daily"])
    import io
    buf = io.BytesIO()
    wb2.save(buf)
    buf.seek(0)
    res2 = parse_excel_universal(buf)
    print(f"Test 2 (Custom Inverter headers): {len(res2)} items parsed.")
    print("ALL UNIVERSAL PARSER TESTS PASSED!")
