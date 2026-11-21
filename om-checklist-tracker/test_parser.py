import re
from openpyxl import load_workbook

def test_parser(filepath):
    wb = load_workbook(filepath, read_only=True, data_only=True)
    sheet = wb["spv"] if "spv" in wb.sheetnames else wb.active

    def _cell_text(v):
        return "" if v is None else str(v).strip()

    def _norm(v):
        return re.sub(r'[\s_]+', ' ', _cell_text(v)).lower()

    col_map = {}
    header_row = None

    CHECKPOINT_HINTS = ["checkpoint", "check point", "point no", "s.no", "sl.no", "sr.no", "sl no", "sr no", "sl", "sr", "s/n", "item no", "no.", "#", "no"]
    DESC_HINTS = ["description", "job description", "observation", "activity", "parameter", "particulars", "check", "item"]
    STANDARD_HINTS = ["standard", "specification", "acceptance", "criteria", "limits", "norm"]
    METHOD_HINTS = ["method", "tool", "instrument", "procedure"]
    FREQ_HINTS = ["frequency", "freq", "period", "interval"]
    SECTION_HINTS = ["section", "category", "system", "subsystem", "area", "group"]

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

    print(f"Header row: {header_row}, col_map: {col_map}")

    # Test row parsing
    items = []
    current_section = ""
    for ri, row in enumerate(sheet.iter_rows(min_row=(header_row or 1) + 1, values_only=True), start=(header_row or 1) + 1):
        row = list(row)
        if not any(v is not None and str(v).strip() for v in row):
            continue

        sec_idx = col_map.get("section", 1)
        cp_idx = col_map.get("checkpoint", 2)
        desc_idx = col_map.get("description", 3)
        std_idx = col_map.get("standard", 4)
        mth_idx = col_map.get("method", 5)
        frq_idx = col_map.get("frequency", 6)

        sec_val = _cell_text(row[sec_idx]) if sec_idx is not None and sec_idx < len(row) else ""
        if sec_val and not sec_val.isdigit():
            current_section = sec_val

        cp_val = _cell_text(row[cp_idx]) if cp_idx is not None and cp_idx < len(row) else ""
        desc_val = _cell_text(row[desc_idx]) if desc_idx is not None and desc_idx < len(row) else ""

        # If cp_val is not numeric or desc_val is empty, check if row matches other column offsets
        if not desc_val or desc_val.lower() in ("job description", "description"):
            continue

        items.append((cp_val, desc_val, current_section))

    print(f"Total items parsed: {len(items)}")
    for it in items[:5]:
        print("  ", it)
    print("...")
    for it in items[-3:]:
        print("  ", it)

if __name__ == "__main__":
    test_parser("Trafo_Checklist_Nov_2026_Populated.xlsx")
