import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
import json
import glob
import re
import random

thin = Border(left=Side(style='thin'), right=Side(style='thin'),
              top=Side(style='thin'), bottom=Side(style='thin'))

with open("data/time_slots.json") as f:
    slots = json.load(f)["time_slots"]

def t2m(x):
    h, m = map(int, x.split(":"))
    return h * 60 + m

slots_sorted = sorted(slots, key=lambda z: t2m(z["start"]))
slot_keys = [f"{s['start']}-{s['end']}" for s in slots_sorted]
slot_dur = {f"{s['start']}-{s['end']}": (t2m(s["end"]) - t2m(s["start"])) / 60.0 for s in slots_sorted}
days = ["Monday","Tuesday","Wednesday","Thursday","Friday"]

def extract_code(x):
    if x is None: return ""
    x = str(x).strip()
    if x == "": return ""
    return x.split()[0].upper()

def split_faculty(fac):
    if fac is None: return []
    fac = str(fac)
    parts = [x.strip() for x in re.split(r"[\\/;,&]| and ", fac) if x.strip() != ""]
    return parts

def load_all_course_info():
    files = glob.glob("data/*.csv")
    fac_map = {}
    p_map = {}
    elective_by_basket = {}
    for f in files:
        try:
            df = pd.read_csv(f)
        except:
            continue
        cols = {c.lower(): c for c in df.columns}
        code_col = cols.get("course_code") or cols.get("course code")
        fac_col = cols.get("faculty")
        ltp_col = cols.get("l-t-p-s-c") or cols.get("l_t_p_s_c") or cols.get("ltp")
        elec_col = cols.get("elective")
        basket_col = cols.get("electivebasket") or cols.get("elective_basket")
        if code_col is None:
            continue
        for _, r in df.iterrows():
            c = str(r.get(code_col, "")).strip()
            if c == "" or c.lower() in {"nan", "new"}:
                continue
            c_up = c.upper()
            fac_raw = str(r.get(fac_col, "")).strip() if fac_col else ""
            facs = split_faculty(fac_raw)
            fac_map[c_up] = facs
            ltp = str(r.get(ltp_col, "")).strip() if ltp_col else ""
            try:
                parts = [float(x) for x in re.split(r"[-:]", ltp) if x != ""]
                P = parts[2] if len(parts) >= 3 else 0.0
            except:
                P = 0.0
            p_map[c_up] = P
            elect = str(r.get(elec_col, "")).strip() if elec_col else ""
            basket = str(r.get(basket_col, "")).strip() if basket_col else ""
            if elect == "1" or elect.lower() == "yes":
                key = basket if basket and basket != "0" else c_up
                elective_by_basket.setdefault(key, []).append(c_up)
    return fac_map, p_map, elective_by_basket
def add_section(cell_value, section):
    cell_value = str(cell_value)
    # prevent duplicate labels
    if "(" in cell_value and cell_value.endswith(")"):
        return cell_value
    return f"{cell_value} ({section})"

course_faculty, course_P, elective_baskets = load_all_course_info()

colors = [
    "FFB3BA","BAE1FF","BAFFC9","FFFFBA","FFD8BA","E3BAFF","D0BAFF","FFCBA4",
    "C7FFD8","B8E1FF","F7FFBA","FFDFBA","E9BAFF","BAFFD9","FFE1BA","BAFFF2",
    "D1FFBA","B2D8F7","F2C2FF","C2FFD8","FFB8E1","D8FFB8","FFE3BA","BAE7FF",
    "E8BAFF","BAFFD6","FFF2BA","DAD7FF","BFFFE1","FFDAB8","E2FFBA","BAF7FF"
]
random.shuffle(colors)
color_map = {}
fill_map = {}
def get_fill(course_code):
    k = course_code.strip().upper()
    if k == "": return None
    if k not in color_map:
        color_map[k] = colors.pop() if colors else "CCCCCC"
    col = color_map[k]
    if col not in fill_map:
        fill_map[col] = PatternFill(start_color=col, end_color=col, fill_type="solid")
    return fill_map[col]

wb_in = openpyxl.load_workbook("Balanced_Timetable_latest.xlsx")
faculty_slots = {}

for sheet in wb_in.sheetnames:
    ws = wb_in[sheet]
    rows = list(ws.values)
    if not rows: continue
    header_index = None
    for i, row in enumerate(rows):
        if row and row[0] == "Day":
            header_index = i
            break
    if header_index is None: continue
    header = rows[header_index]
    col_map = {name: idx for idx, name in enumerate(header) if name in slot_keys}
    merged_ranges = list(ws.merged_cells.ranges)
    merged_lookup = {}
    for rng in merged_ranges:
        minr, minc, maxr, maxc = rng.min_row, rng.min_col, rng.max_row, rng.max_col
        for c in range(minc, maxc + 1):
            merged_lookup[(minr, c)] = (minc, maxc)
    cleaned = {d: {s: "" for s in slot_keys} for d in days}
    for r in range(header_index + 1, ws.max_row + 1):
        day_val = ws.cell(r, 1).value
        if day_val not in days: continue
        d = day_val
        for sk, col in col_map.items():
            c = col + 1
            key = (r, c)
            if key in merged_lookup:
                minc, maxc = merged_lookup[key]
                val = ws.cell(r, minc).value
                for cc in range(minc, maxc + 1):
                    hdr = header[cc - 1]
                    if hdr in slot_keys:
                        cleaned[d][hdr] = val if val is not None else ""
            else:
                val = ws.cell(r, c).value
                cleaned[d][sk] = val if val is not None else ""
    for d in days:
        for idx, sk in enumerate(slot_keys):
            cell = cleaned[d][sk]
            code = extract_code(cell)
            if code == "": continue
            P = course_P.get(code, 0.0)
            # If this slot is already part of a multi-slot merged entry, do NOT expand it again
            if idx > 0 and cleaned[d][slot_keys[idx-1]] == cell:
                continue

            # PATCHED LAB EXPANSION (only fills empty slots, avoids infinite spreading)
            if P >= 2.0 or "LAB" in str(cell).upper():
                total = 0.0
                k = idx
                while k < len(slot_keys) and total < 2.0:
                    target = slot_keys[k]

        # Only allow expansion if the slot is EMPTY
                    if cleaned[d][target] == "":
                        cleaned[d][target] = cell
                        total += slot_dur[target]
                        k += 1
                    else:
                        break

    for d in days:
        for sk in slot_keys:
            cell = cleaned[d][sk]
            if cell is None or str(cell).strip() == "": continue
            code = extract_code(cell)
            if code.startswith("ELECTIVE"):
                m = re.search(r"ELECTIVE\s*([0-9]+)", code)

    # always add section name
                val_with_section = add_section(cell, sheet)

                if m:
                    b = m.group(1)
                    reps = elective_baskets.get(b, [])
                    for c_up in reps:
                        facs = course_faculty.get(c_up, [])
                        for fac in facs:
                            if fac == "":
                                continue
                            faculty_slots.setdefault(fac, {dd: {s: "" for s in slot_keys} for dd in days})
                            faculty_slots[fac][d][sk] = val_with_section
                    continue

    # fallback block (your else part)
                for key, reps in elective_baskets.items():
                    for c_up in reps:
                        facs = course_faculty.get(c_up, [])
                        for fac in facs:
                            if fac == "":
                                continue
                            faculty_slots.setdefault(fac, {dd: {s: "" for s in slot_keys} for dd in days})
                            faculty_slots[fac][d][sk] = val_with_section
                continue

            facs = course_faculty.get(code, [])
            if facs:
                for fac in facs:
                    if fac == "": continue
                    faculty_slots.setdefault(fac, {dd: {s: "" for s in slot_keys} for dd in days})
                    faculty_slots[fac][d][sk] = cell

def safe_title(x):
    return re.sub(r'[:\\/*?\[\]]', '_', x)[:30]

wb_out = Workbook()
first = True

for fac, table in faculty_slots.items():
    title = safe_title(fac if fac else "Unknown")
    if first:
        ws = wb_out.active
        ws.title = title
        first = False
    else:
        ws = wb_out.create_sheet(title)
    ws.append(["Day"] + slot_keys)
    for d in days:
        ws.append([d] + [table[d][s] for s in slot_keys])
    for r in range(2, ws.max_row + 1):
        c = 2
        while c <= ws.max_column:
            val = ws.cell(r, c).value
            if not val:
                c += 1
                continue
            start = c
            end = c
            while end + 1 <= ws.max_column and ws.cell(r, end + 1).value == val:
                end += 1
            code = extract_code(val)
            fill = get_fill(code) if code else None
            if end > start:
                ws.merge_cells(start_row=r, start_column=start, end_row=r, end_column=end)
            top_left = ws.cell(r, start)
            if fill:
                top_left.fill = fill
            for cc in range(start, end + 1):
                cell = ws.cell(r, cc)
                cell.border = thin
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                if cell.row == 1:
                    cell.font = Font(bold=True)
                if fill:
                    cell.fill = fill
            c = end + 1
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.row == 1:
                cell.font = Font(bold=True)
            cell.border = thin
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                max_len = max(max_len, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = min(max_len + 2, 60)

wb_out.save("Faculty_Timetable.xlsx")
print("Faculty_Timetable.xlsx generated successfully.")
