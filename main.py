import pandas as pd
import json
import random
with open("time_slots.json") as f:
    slots = json.load(f)["time_slots"]

slot_keys = [f"{slot['start'].strip()}-{slot['end'].strip()}" for slot in slots]

def slot_duration(slot):
    start, end = slot.split("-")
    h1, m1 = map(int, start.split(":"))
    h2, m2 = map(int, end.split(":"))
    return (h2 + m2 / 60) - (h1 + m1 / 60)

slot_durations = {s: slot_duration(s) for s in slot_keys}

courses = pd.read_csv("courses.csv").to_dict(orient="records")
rooms_df = pd.read_csv("rooms.csv")  # must have columns: Room_ID, Type
classrooms = rooms_df[rooms_df["Type"].str.lower() == "classroom"]["Room_ID"].tolist()
labs = rooms_df[rooms_df["Type"].str.lower() == "lab"]["Room_ID"].tolist()

days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
excluded_slots = ["07:30-09:00", "13:15-14:00", "17:30-18:30"]
MAX_ATTEMPTS = 10

def get_free_blocks(timetable, day):
    free_blocks = []
    block = []
    for slot in slot_keys:
        if timetable.at[day, slot] == "" and slot not in excluded_slots:
            block.append(slot)
        else:
            if block:
                free_blocks.append(block)
                block = []
    if block:
        free_blocks.append(block)
    return free_blocks

def allocate_session(timetable, lecturer_busy, day, faculty, code, duration_hours, session_type="L"):
    free_blocks = get_free_blocks(timetable, day)
    for block in free_blocks:
        total = sum(slot_durations[s] for s in block)
        if total >= duration_hours:
            slots_to_use = []
            dur_accum = 0
            for s in block:
                slots_to_use.append(s)
                dur_accum += slot_durations[s]
                if dur_accum >= duration_hours:
                    break
            
            # Assign room based on session type
            if session_type == "P":
                if not labs:
                    print(f"No labs available for {code}")
                    return False
                room = random.choice(labs)
            else:  # L or T
                if not classrooms:
                    print(f"No classrooms available for {code}")
                    return False
                room = random.choice(classrooms)

            for s in slots_to_use:
                if session_type == "L":
                    timetable.at[day, s] = f"{code} ({room})"
                elif session_type == "T":
                    timetable.at[day, s] = f"{code}T ({room})"
                elif session_type == "P":
                    timetable.at[day, s] = f"{code} (Lab-{room})"

            if faculty:
                lecturer_busy[day].append(faculty)
            return True
    return False

def generate_timetable(courses_to_allocate, filename):
    timetable = pd.DataFrame("", index=days, columns=slot_keys)
    lecturer_busy = {day: [] for day in days}

    electives = [c for c in courses_to_allocate if str(c.get("Elective", 0)) == "1"]
    non_electives = [c for c in courses_to_allocate if str(c.get("Elective", 0)) != "1"]

    if electives:
        chosen_elective = random.choice(electives)
        elective_course = {
            "Course_Code": "Elective",
            "Faculty": chosen_elective.get("Faculty", ""),
            "L-T-P-S-C": chosen_elective["L-T-P-S-C"]
        }
        non_electives.append(elective_course)

    for course in non_electives:
        faculty = str(course.get("Faculty", "")).strip()
        code = str(course["Course_Code"]).strip()

        try:
            L, T, P, S, C = map(int, [x.strip() for x in course["L-T-P-S-C"].split("-")])
        except:
            continue

        lecture_hours_remaining = L
        attempts = 0
        while lecture_hours_remaining > 0 and attempts < MAX_ATTEMPTS:
            attempts += 1
            for day in days:
                if lecture_hours_remaining <= 0 or (faculty and faculty in lecturer_busy[day]):
                    continue
                alloc_hours = min(1.5, lecture_hours_remaining)
                if allocate_session(timetable, lecturer_busy, day, faculty, code, alloc_hours, "L"):
                    lecture_hours_remaining -= alloc_hours
                    break
        if lecture_hours_remaining > 0:
            print(f"Warning: Could not fully allocate lectures for {code}")

        tutorial_hours_remaining = T
        attempts = 0
        while tutorial_hours_remaining > 0 and attempts < MAX_ATTEMPTS:
            attempts += 1
            for day in days:
                if tutorial_hours_remaining <= 0 or (faculty and faculty in lecturer_busy[day]):
                    continue
                if allocate_session(timetable, lecturer_busy, day, faculty, code, 1, "T"):
                    tutorial_hours_remaining -= 1
                    break
        if tutorial_hours_remaining > 0:
            print(f"Warning: Could not fully allocate tutorials for {code}")

        practical_hours_remaining = P
        attempts = 0
        while practical_hours_remaining > 0 and attempts < MAX_ATTEMPTS:
            attempts += 1
            for day in days:
                if practical_hours_remaining <= 0 or (faculty and faculty in lecturer_busy[day]):
                    continue
                alloc_hours = min(2, practical_hours_remaining) if practical_hours_remaining >= 2 else practical_hours_remaining
                if allocate_session(timetable, lecturer_busy, day, faculty, code, alloc_hours, "P"):
                    practical_hours_remaining -= alloc_hours
                    break
        if practical_hours_remaining > 0:
            print(f"Warning: Could not fully allocate practicals for {code}")

    for day in days:
        for slot in excluded_slots:
            if slot in timetable.columns:
                timetable.at[day, slot] = ""

    timetable.to_excel(filename, index=True)
    print(f"Saved timetable to {filename}")

courses_first_half = [c for c in courses if str(c.get("Semester_Half")).strip() in ["1", "0"]]
courses_second_half = [c for c in courses if str(c.get("Semester_Half")).strip() in ["2", "0"]]

generate_timetable(courses_first_half, "timetable_first_half.xlsx")
generate_timetable(courses_second_half, "timetable_second_half.xlsx")
