# 📘 Automated Timetable + Exam Scheduler + Faculty Timetable Generator  
### **Complete Python Automation Suite for IIIT Timetables & Exams**

This repository contains a fully automated system for generating:

✔ **Class Timetables** (with room allocation, elective sync, merging, color coding)  
✔ **Exam Schedules** (merged across departments, room-based allocation, invigilation)  
✔ **Faculty-wise Timetables** (per-faculty personalized grids based on generated timetable)

It consists of **three major Python modules**:

timetable.py → Generates balanced class timetables
exam.py → Generates student-wise exam schedules with room allocation
faculty.py → Generates per-faculty timetable sheets

All outputs are exported in **Excel** format with automatic formatting, merging, coloring, and legends.

---

# 🚀 Features

## ✅ 1. **Balanced Class Timetable Generator (`timetable.py`)**
- Automatic slot allocation using:
  - L/T/P hours extracted from L-T-P-S-C  
  - Faculty availability constraints  
  - Room capacity matching  
  - Lab vs. classroom separation  
  - Break/excluded slots  
- Multi-section support:
  - **CSEA I / CSEB I / CSEA III / CSEB III / CSE-V / DSAI / ECE / Semester 7**
- Elective basket handling with **synchronized elective slots across branches**
- Auto-coloring and merging of identical adjacent cells
- Automatically generates legends for each section
- Produces: Balanced_Timetable_latest.xlsx

---

## ✅ 2. **Exam Scheduler (`exam.py`)**
- Handles merged departmental exams
- Constraints enforced:
  - Max **1 exam per group/day**
  - Max **4 global exams/day**
- Room allocation:
  - Labs, libraries excluded  
  - Normal rooms used before halls  
  - Half-capacity seating rule applied  
- Electives handled semester-wise (morning/afternoon batch logic)
- Auto invigilator assignment (capacity-based: 1 or 2 invigilators)
- Exports: final_exam_schedule.xlsx

Includes:
- Merged exam rows
- Grid-view day/slot table
- Course legend
- Full styling, alignment, alternating color rows

---

## ✅ 3. **Faculty Timetable Generator (`faculty.py`)**
- Reads generated timetable and all course CSVs
- Extracts faculty-wise:
  - Day  
  - Slot  
  - Room  
  - Course  
  - Section  
- Creates a clean **grid timetable for each faculty member**
- Includes index sheet listing faculty and number of classes
- Produces: Faculty_Timetable_Grid.xlsx

---

# 📂 Project Structure

├── timetable.py
├── exam.py
├── faculty.py
├── data/
│   ├── rooms.csv
│   ├── faculty.csv
│   ├── time_slots.json
│   └── courses*.csv
├──docs/
│   ├── SoftwareSages_DPR.pdf
│   └── SoftwareSages_DPR.tex
├── Balanced_Timetable_latest.xlsx 
├── final_exam_schedule.xlsx 
└── Faculty_Timetable_Grid_*.xlsx

---

# 🧩 How Everything Works

## **Timetable Generation Flow**
1. Parse time slots + normalize durations  
2. Load course CSVs and split into first/second semester half  
3. Allocate L/T/P hours with:
   - Synchronized elective slots  
   - Contiguous blocks search  
   - Excluded slot avoidance  
4. Allocate rooms:
   - Capacity-based  
   - Lab/classroom enforcement  
   - Same course → same room reuse  
5. Auto merge identical cells in Excel  
6. Apply per-course colors  
7. Add legends & export

---

## **Exam Scheduling Flow**
1. Read department exam CSVs  
2. Build elective pools per semester  
3. Morning/afternoon elective split  
4. Allocate rooms using:
   - Best-fit order  
   - Half-capacity  
   - Hall-last system  
5. Schedule regular courses with constraints  
6. Assign invigilators  
7. Export formatted Excel

---

## **Faculty Timetable Extraction Flow**
1. Read generated class timetable  
2. For each entry: extract  
   - Course code  
   - Room  
   - Day  
   - Slot  
3. Match with CSV course info to get faculty  
4. Group by faculty  
5. Generate individual colored sheets

---

# ⚙️ Installation

### **Requirements**
1. Python 3.10+
2. pandas
3. openpyxl


## 🏃 Usage

---

### **1. Generate Class Timetable**

  python timetable.py

Ouput:
  Balanced_Timetable_latest.xlsx

### **2. Generate Exam Timetable**

  python exam.py

Output:
  final_exam_schedule.xlsx


### **3. Generate Faculty Timetable**

  python faculty.py

Output:
  Faculty_Timetable_Grid.xlsx

## Team Members (Software Sages)
- Nikunj Srivastav (24BCS087)
- Sudhanshu Baberwal (24BCS147)
- Thejas Gowda U M (24BCS157)
- Shaik Moiz (24BCS133)