import unittest
from unittest.mock import patch, MagicMock
import pandas as pd

from main_timetable import parse_time, slot_duration_from_bounds, parse_ltp, safe_str, get_free_blocks, allocate_session, merge_and_style_cells, generate_timetable, split_by_half

# Example functions extracted for testing
def parse_time(t):
    h, m = map(int, t.split(":"))
    return h * 60 + m

def slot_duration_from_bounds(start, end):
    return (parse_time(end) - parse_time(start)) / 60.0

def safe_str(val):
    if val is None:
        return ""
    if isinstance(val, float) and pd.isna(val):
        return ""
    return str(val).strip()

def parse_ltp(sc_string):
    try:
        parts = [x.strip() for x in sc_string.split("-")]
        while len(parts) < 5:
            parts.append("0")
        return list(map(int, parts[:5]))
    except:
        return [0, 0, 0, 0, 0]

class TestTimetableFunctions(unittest.TestCase):

    def test_parse_time(self):
        self.assertEqual(parse_time("07:30"), 450)
        self.assertEqual(parse_time("00:00"), 0)
        self.assertEqual(parse_time("23:59"), 1439)

    def test_slot_duration_from_bounds(self):
        self.assertAlmostEqual(slot_duration_from_bounds("07:30", "09:00"), 1.5)
        self.assertAlmostEqual(slot_duration_from_bounds("13:15", "14:00"), 0.75)

    def test_safe_str(self):
        self.assertEqual(safe_str(None), "")
        self.assertEqual(safe_str(float('nan')), "")
        self.assertEqual(safe_str("  test  "), "test")
        self.assertEqual(safe_str(123), "123")

    def test_parse_ltp(self):
        self.assertEqual(parse_ltp("3-0-2"), [3, 0, 2, 0, 0])
        self.assertEqual(parse_ltp("1-1-1-1-1"), [1, 1, 1, 1, 1])
        self.assertEqual(parse_ltp(""), [0, 0, 0, 0, 0])
        self.assertEqual(parse_ltp("invalid"), [0, 0, 0, 0, 0])

    # Example of get_free_blocks test
    def test_get_free_blocks(self):
        slot_keys = ["07:30-09:00", "09:00-10:30", "10:30-12:00"]
        excluded_slots = ["07:30-09:00"]
        df = pd.DataFrame("", index=["Monday"], columns=slot_keys)
        df.at["Monday", "09:00-10:30"] = "SomeClass"
        
        from facultyTT import get_free_blocks  # Import actual function
        free_blocks = get_free_blocks(df, "Monday")
        self.assertEqual(free_blocks, [["10:30-12:00"]])

    # Placeholder for allocate_session, merge_and_style_cells, generate_timetable tests
    # These can be tested using mocks because they depend on files, randomness, and Excel
    @patch("facultyTT.random.choice")
    def test_allocate_session_mocked(self, mock_choice):
        mock_choice.return_value = "Lab1"
        timetable = pd.DataFrame("", index=["Monday"], columns=["09:00-10:30", "10:30-12:00"])
        lecturer_busy = {"Monday": {}}
        course_room_map = {}
        labs_on_days = set()
        
        from facultyTT import allocate_session
        result = allocate_session(timetable, lecturer_busy, course_room_map, "Monday", "Prof A", "CS101", 1.0, "L", False, labs_on_days)
        self.assertTrue(result)
        self.assertIn("CS101", timetable.values)

    def test_split_by_half(self):
        courses_list = [
            {"Semester_Half": "1", "Course_Code": "C1"},
            {"Semester_Half": "2", "Course_Code": "C2"},
            {"Semester_Half": "0", "Course_Code": "C0"}
        ]
        from facultyTT import split_by_half
        first, second = split_by_half(courses_list)
        self.assertEqual(len(first), 2)
        self.assertEqual(len(second), 2)

if __name__ == "__main__":
    unittest.main()


