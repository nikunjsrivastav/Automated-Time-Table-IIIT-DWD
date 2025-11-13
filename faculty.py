import pandas as pd
import re, os, random
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

# --- setup ---
random.seed(123)
thin = Border(left=Side(style='thin'),
              right=Side(style='thin'),
              top=Side(style='thin'),
              bottom=Side(style='thin'))
days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
excluded = ["07:30-09:00", "10:30-10:45", "13:15-14:00", "15:30-15:40"]

colors = [
    "FFB3BA","BAE1FF","BAFFC9","FFFFBA","FFD8BA","E3BAFF","D0BAFF","FFCBA4",
    "C7FFD8","B8E1FF","F7FFBA","FFDFBA","E9BAFF","BAFFD9","FFE1BA","BAFFF2",
    "D1FFBA","B2D8F7","F2C2FF","C2FFD8","FFB8E1","D8FFB8","FFE3BA","BAE7FF",
    "E8BAFF","BAFFD6","FFF2BA","DAD7FF","BFFFE1","FFDAB8","E2FFBA","BAF7FF"
]

# --- helper ---
def extract_info(val):
    if not isinstance(val, str) or val.strip() == "":
        return None, None
    m = re.match(r"([A-Z]{1,5}\d{0,3})(?:.*\((.*?)\))?", val.strip())
    if not m:
        return None, None
    return m.group(1).strip(), m.group(2).strip() if m.group(2) else ""

def clean_faculty(name):
    if not isinstance(name, str): return ""
    return [n.strip() for n in name.replace("/", ",").split(",") if n.strip()]

def build_faculty_data(timetable_path, csvs):
    all_courses = []
    for csv in csvs:
        if os.path.exists(csv):
            df = pd.read_csv(csv)
            df["Section"] = os.path.basename(csv).replace(".csv", "")
            all_courses.append(df)
    if not all_courses:
        raise FileNotFoundError("No CSV files found.")
    df_all = pd.concat(all_courses, ignore_index=True)
    df_all["Faculty"] = df_all["Faculty"].fillna("").astype(str)
    df_all["Course_Code"] = df_all["Course_Code"].fillna("").astype(str)
    df_all["Course_Code"] = df_all["Course_Code"].str.upper().str.strip()

    xl = pd.ExcelFile(timetable_path)
    fac_map = {}

    for sheet in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet, header=None)
        if df.empty:
            continue
        idx = df.index[df.iloc[:,0] == "Day"].tolist()
        if not idx:
            continue
        row = idx[0]
        slots = df.iloc[row, 1:].dropna().tolist()

        for i in range(row+1, len(df)):
            day = df.iat[i, 0]
            if day not in days:
                continue
            for j, slot in enumerate(slots, start=1):
                val = df.iat[i, j]
                code, room = extract_info(val)
                if not code: continue
                course_info = df_all[df_all["Course_Code"].str.upper() == code.upper()]
                if course_info.empty: continue
                for _, c in course_info.iterrows():
                    for fac in clean_faculty(c["Faculty"]):
                        fac_map.setdefault(fac, []).append({
                            "day": day,
                            "slot": slot,
                            "code": code,
                            "title": c.get("Course_Title",""),
                            "room": room,
                            "section": c.get("Section","")
                        })
    return fac_map

def create_faculty_timetables(fac_map, slot_order):
    wb = Workbook()
    ws_index = wb.active
    ws_index.title = "Faculty Index"
    ws_index.append(["Faculty", "Total Classes"])
    ws_index.cell(1,1).font = Font(bold=True)
    ws_index.cell(1,2).font = Font(bold=True)

    for idx, (fac, items) in enumerate(sorted(fac_map.items()), start=2):
        ws_index.append([fac, len(items)])
        ws_index.cell(idx,1).border = thin
        ws_index.cell(idx,2).border = thin

    color_avail = colors.copy()
    random.shuffle(color_avail)
    color_map = {}

    for fac, entries in fac_map.items():
        ws = wb.create_sheet(fac[:31])
        ws.append(["Day"] + slot_order)
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin
            cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

        tt = pd.DataFrame("", index=days, columns=slot_order)
        for e in entries:
            tt.at[e["day"], e["slot"]] = f"{e['code']} ({e['room']})"

        for r_idx, d in enumerate(days, start=2):
            row = [d] + [tt.at[d, s] for s in slot_order]
            ws.append(row)
            for c in range(1, len(row)+1):
                cell = ws.cell(r_idx, c)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = thin
                val = str(cell.value).strip()
                if val:
                    code = val.split()[0]
                    if code not in color_map:
                        color_map[code] = color_avail.pop() if color_avail else "CCCCCC"
                    color = color_map[code]
                    cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 15

    name = f"Faculty_Timetable_Grid_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(name)
    print("✅ Faculty timetables (grid format) saved as", name)


if __name__ == "__main__":
    # auto detect latest timetable
    latest = max([f for f in os.listdir() if f.startswith("Balanced_Timetable_") and f.endswith(".xlsx")], key=os.path.getctime)
    csvs = [
        "data/coursesCSEA-I.csv", "data/coursesCSEB-I.csv", "data/coursesCSEA-III.csv", "data/coursesCSEB-III.csv",
        "data/coursesCSE-V.csv", "data/coursesDSAI-III.csv", "data/coursesECE-III.csv", "data/courses7.csv",
        "data/coursesDSAI-I.csv", "data/coursesDSAI-V.csv", "data/coursesECE-I.csv", "data/coursesECE-V.csv"
    ]
    xl = pd.ExcelFile(latest)
    df_sample = pd.read_excel(xl, sheet_name=xl.sheet_names[0], header=None)
    slot_row_idx = df_sample.index[df_sample.iloc[:,0] == "Day"].tolist()[0]
    slot_order = df_sample.iloc[slot_row_idx, 1:].dropna().tolist()

    fac_map = build_faculty_data(latest, csvs)
    create_faculty_timetables(fac_map, slot_order)