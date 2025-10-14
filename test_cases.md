# 🧪 Test Cases — Automated Timetable Generator

This document lists all test cases for verifying each function in the timetable generation system.  
Each table follows the format:

| Test Case Input | Function | Description | Expected Output |
|-----------------|-----------|--------------|-----------------|

---

## **1. `parse_time(t)`**

| Test Case Input | Function | Description | Expected Output |
|-----------------|-----------|--------------|-----------------|
| `"07:30"` | parse_time | Convert valid morning time to minutes | `450` |
| `"00:00"` | parse_time | Midnight edge case | `0` |
| `"23:59"` | parse_time | End of day edge case | `1439` |
| `"09:00"` | parse_time | Standard slot start time | `540` |
| `"invalid"` | parse_time | Invalid time format | Should raise or handle error gracefully |

---

## **2. `slot_duration_from_bounds(start, end)`**

| Test Case Input | Function | Description | Expected Output |
|-----------------|-----------|--------------|-----------------|
| `"07:30", "09:00"` | slot_duration_from_bounds | Valid 1.5 hour duration | `1.5` |
| `"09:00", "10:00"` | slot_duration_from_bounds | One-hour duration | `1.0` |
| `"17:30", "18:30"` | slot_duration_from_bounds | Evening slot | `1.0` |
| `"14:00", "14:00"` | slot_duration_from_bounds | Zero-length duration | `0.0` |
| `"18:00", "17:00"` | slot_duration_from_bounds | End before start edge | Should raise error or return negative duration |

---

## **3. `safe_str(val)`**

| Test Case Input | Function | Description | Expected Output |
|-----------------|-----------|--------------|-----------------|
| `"CSE101"` | safe_str | Normal string | `"CSE101"` |
| `None` | safe_str | Handles `None` safely | `""` |
| `float('nan')` | safe_str | Handles NaN safely | `""` |
| `"  test  "` | safe_str | Trims extra spaces | `"test"` |
| `1234` | safe_str | Converts non-string input | `"1234"` |

---

## **4. `parse_ltp(sc_string)`**

| Test Case Input | Function | Description | Expected Output |
|-----------------|-----------|--------------|-----------------|
| `"3-1-2-0-6"` | parse_ltp | Normal case | `[3,1,2,0,6]` |
| `"4-0-0"` | parse_ltp | Missing values filled with zeros | `[4,0,0,0,0]` |
| `"2 - 1 - 0"` | parse_ltp | Handles extra spaces | `[2,1,0,0,0]` |
| `""` | parse_ltp | Empty input string | `[0,0,0,0,0]` |
| `"a-b-c"` | parse_ltp | Invalid non-numeric input | `[0,0,0,0,0]` |

---

## **5. `get_free_blocks(timetable, day)`**

| Test Case Input | Function | Description | Expected Output |
|-----------------|-----------|--------------|-----------------|
| Empty timetable for Monday | get_free_blocks | Detects all slots as free except excluded | List of all available slots |
| Partially filled Tuesday | get_free_blocks | Splits free slots before & after occupied ones | Multiple free blocks returned |
| Day with all excluded slots filled | get_free_blocks | Ensures excluded slots not considered | Free blocks ignore exclusions |
| Full Thursday | get_free_blocks | No empty slots available | Empty list |

---

## **6. `allocate_session(...)`**

| Test Case Input | Function | Description | Expected Output |
|-----------------|-----------|--------------|-----------------|
| L-type session, 1.5 hr, empty timetable | allocate_session | Normal successful allocation | Returns `True`, slot filled |
| Same faculty busy same day/time | allocate_session | Faculty conflict | Returns `False`, no allocation |
| P-type session, no labs available | allocate_session | Lab unavailable | Returns `False`, error message logged |
| Slot overlaps excluded slot | allocate_session | Invalid slot allocation skipped | Returns `False` |
| P-type session with labs available | allocate_session | Lab randomly chosen | Returns `True` with lab room assigned |
| Course already has assigned room | allocate_session | Ensures room consistency | Reuses same room, returns `True` |

---

## **7. `merge_and_style_cells(filename)`**

| Test Case Input | Function | Description | Expected Output |
|-----------------|-----------|--------------|-----------------|
| Adjacent identical course names | merge_and_style_cells | Merges consecutive cells | Single merged cell |
| Different course values | merge_and_style_cells | No merge between different values | Cells remain separate |
| Randomized color pool | merge_and_style_cells | Each course has unique color | Visually distinct cells |
| Empty cells | merge_and_style_cells | Skips merging/styling | No changes applied |
| Long text values | merge_and_style_cells | Tests wrapping and alignment | Text wraps properly inside cell |

---

## **8. `generate_timetable(courses_to_allocate, filename)`**

| Test Case Input | Function | Description | Expected Output |
|-----------------|-----------|--------------|-----------------|
| Valid course list (CSE) | generate_timetable | Standard timetable generation | Creates formatted Excel file |
| Only elective courses | generate_timetable | Tests elective handling | Adds "Elective" placeholder course |
| Too many courses | generate_timetable | Stress test with insufficient slots | Partial allocation, no crash |
| Missing faculty name | generate_timetable | Tests optional faculty handling | Allocates ignoring busy map |
| No lab rooms | generate_timetable | Tests lab failure case | Skips lab allocation, handles gracefully |
| Overbooked timetable | generate_timetable | Checks slot exhaustion handling | Courses left unallocated |
| Output validation | generate_timetable | Checks Excel output styling | Formatted file with merged cells |

---

## **9. `split_by_half(courses_list)`**

| Test Case Input | Function | Description | Expected Output |
|-----------------|-----------|--------------|-----------------|
| Courses with Semester_Half=1,2 | split_by_half | Normal splitting | Two separate lists |
| Courses with Semester_Half=0 | split_by_half | Shared across both halves | Appears in both outputs |
| Missing Semester_Half key | split_by_half | Default behavior | Added to first half |
| Empty list | split_by_half | Edge case | Returns two empty lists |

---

## **10. Integrated Tests**

| Test Case Input | Function | Description | Expected Output |
|-----------------|-----------|--------------|-----------------|
| All CSV files (CSE, ECE, DSAI) present | Full Run | Normal end-to-end execution | Generates 6 `.xlsx` files |
| One CSV missing | Full Run | Missing dataset | Raises FileNotFoundError or handled gracefully |
| Duplicate course codes | Full Run | Check room consistency | Same room reused for duplicates |
| Shared faculty across courses | Full Run | Checks conflict management | No overlapping slots |
| Repeated run with same seed | Full Run | Deterministic behavior | Identical results across runs |

---

## **11. Performance & Robustness Tests**

| Test Case Input | Function | Description | Expected Output |
|-----------------|-----------|--------------|-----------------|
| 100+ courses | Full Program | Stress test | Executes under 5s, no crash |
| Corrupted `time_slots.json` | Full Program | JSON input error handling | Graceful failure with message |
| Empty `courses.csv` | Full Program | No courses to schedule | Generates empty timetable file |
| Very large room list | Full Program | Tests scalability of room allocation | Handles without slowdown |

---

## **12. Output Validation**

| Test Case Input | Function | Description | Expected Output |
|-----------------|-----------|--------------|-----------------|
| Final Excel outputs | Full Program | Verify generated files | Six `.xlsx` files created |
| Course info appended | Full Program | Verify appended table at end | Course info table formatted with headers |
| Color & border styling | Full Program | Visual format consistency | Colored, bordered, and auto-width columns |
| Timestamped filenames | Full Program | Verify filename pattern | Files saved as `<timestamp>_timetable_*.xlsx` |

---
### **Faculty Timetable Generator (main_facultytt.py)**

| Test Case Input | Function | Description | Expected Output |
|-----------------|----------|-------------|-----------------|
| Valid `time_slots.json` and one or more `courses*.csv` files containing `Faculty` column | `load_time_slots()` / `build_slot_keys()` | Load and normalize time slots from JSON and build `slot_keys` list | `slot_keys` list created (e.g. `["09:00-10:00", ...]`) |
| Missing or malformed `time_slots.json` | `load_time_slots()` | JSON file dependency and parsing error handling | Raises `FileNotFoundError` or `JSONDecodeError` |
| Directory contains no `courses*.csv` files or CSVs missing `Faculty` column | `discover_course_csvs()` / `load_and_concat_courses()` | Discover course CSVs, read them, and concat into single DataFrame | Raises `ValueError: No valid course CSVs found with 'Faculty' column.` |
| Mixed branch CSVs with `BranchFile` assignment | `load_and_concat_courses()` | Ensures branch filename is preserved when loading per-branch CSVs | Combined DataFrame with `BranchFile` column |
| DataFrame with `Faculty` values present | `group_by_faculty()` | Group courses by faculty for per-faculty schedule generation | `faculty_groups` produced (iterable of (faculty, df) pairs) |
| Sheet title contains invalid Excel characters like `/:*?[]` | `sanitize_sheet_name()` | Replace invalid characters and trim to safe length for Excel | Returns sanitized sheet name (underscores, ≤30 chars) |
| Worksheet with variable cell content lengths | `auto_adjust_column_widths()` | Compute and set column widths to fit contents within limit | Columns auto-sized (max width limit applied) |
| Faculty group iteration and sheet creation | `create_faculty_sheets()` | For each faculty, create a sheet and write header row ("Day" + slot keys) | One worksheet per faculty with correct header row |
| Random placement of courses into timetable slots | `assign_random_slots()` | Prototype logic that assigns each course to a random day+slot | Timetable dict filled with course labels (no exceptions) |
| Same course assigned multiple times in a sheet | `assign_random_slots()` | Ensure repeated course occurrences are allowed and handled | Course appears in multiple (random) cells without error |
| Color mapping for course codes and cell styling | `apply_cell_styling()` | Assign consistent color per course; apply borders, alignment, bold font, wrapping | Cells filled and styled; same course shares same fill color |
| Exhausted color pool (more unique courses than preset colors) | `apply_cell_styling()` | Fallback color generation for extra courses | Generates RGB hex colors and applies them consistently |
| Large number of courses for a single faculty (10+) | Full per-faculty generation | Tests scalability of sheet creation, assignment, and styling | Sheet created successfully; all courses assigned; no crash |
| Workbook save step | `save_workbook()` | Persist the final `Workbook` to disk with a stable filename | `facultyTT.xlsx` created on disk |
| Reproducible results when seed set | `random.seed(...)` + overall flow | Deterministic behavior when seed is set before assignment | Same `facultyTT.xlsx` content across runs with same seed |


### ✅ **Expected Deliverables**
- `timetable_first_halfCSE.xlsx`  
- `timetable_second_halfCSE.xlsx`  
- `timetable_first_halfECE.xlsx`  
- `timetable_second_halfECE.xlsx`  
- `timetable_first_halfDSAI.xlsx`  
- `timetable_second_halfDSAI.xlsx`
- `facultyTT.xlsx`

---

### 🏁 **Conclusion**
This test plan ensures that each part of the timetable generation code works correctly, efficiently, and produces visually consistent Excel files while handling invalid data gracefully.

---
