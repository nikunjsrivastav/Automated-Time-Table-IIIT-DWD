import pandas as pd
import json
import random
with open("time_slots.json") as f:
    slots = json.load(f)["time_slots"]
courses = pd.read_csv("courses.csv").to_dict(orient="records")
rooms = pd.read_csv("room.csv")["room_name"].tolist()

days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
timetable = pd.DataFrame(index=[slot["start"] + "-" + slot["end"] for slot in slots],
                         columns=days)
lecturer_busy = {day: [] for day in days}
group_busy = {day: [] for day in days}
for day in days:
    free_rooms = rooms.copy()
    for slot in slots:
        slot_key = slot["start"] + "-" + slot["end"]
        random.shuffle(courses)  # Randomly pick courses
        for course in courses:
            if (course["lecturer"] not in lecturer_busy[day] and
                course["group"] not in group_busy[day] and
                free_rooms):

                room_assigned = free_rooms.pop(0)  # Assign first free room
                timetable.at[slot_key, day] = f"{course['course_name']} ({room_assigned})"
                lecturer_busy[day].append(course["lecturer"])
                group_busy[day].append(course["group"])
                break  # Move to next slot

# Save to Excel
timetable.to_excel("generated_timetable.xlsx")
