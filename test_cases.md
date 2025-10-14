| Test case input | Function | Description | Expected output |
|"3-1-2-0-6" | parse_ltp | valid L-T-P-S-C string | [3,1,2,0,6] |
|"2-0-2" | parse_ltp | short parts -> pad to 5 | [2,0,2,0,0] |
|"abc" | parse_ltp | invalid string | [0,0,0,0,0] |
|"abc-def"| parse_ltp | invalid string | [0,0,0,0,0] |
|("09:00","10:30") | slot_duration_from_bounds | duration calc | 1.5 |
|empty timetable for Monday | get_free_blocks | detect continuous free blocks | list of slot groups |
|No classrooms (empty rooms.csv) | allocate_session | should return False & print message | False & printed "No classrooms available for ..." |
