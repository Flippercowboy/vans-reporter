# Configuration for Vans Department Reporting App

# Monday.com Configuration
BOARD_ID = 1215254769  # Project Tracker
VANS_DEPARTMENT_ID = 17  # Department column value for "Vans"

# Monday.com API Token
# Get your API token from: https://monday.com/developers/v2#authentication-section
# Or set the MONDAY_API_TOKEN environment variable
MONDAY_API_TOKEN = "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjYxNTM1OTA5NSwiYWFpIjoxMSwidWlkIjo0NDIyMTU2MCwiaWFkIjoiMjAyNi0wMi0wMlQxNjozNjoxMi4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6MTcyODIwNjksInJnbiI6ImV1YzEifQ.uPf77jiGt6kjg8yRQOFv8IAr89uoy9RH1szeOtHsOQQ"

# Column IDs
COLUMN_NAME = "name"
COLUMN_STATUS = "status2"
COLUMN_PEOPLE = "people4"
COLUMN_DEPARTMENT = "single_select5"
COLUMN_FILMING_DATE_1 = "date5"
COLUMN_FILMING_DATE_2 = "date14"
COLUMN_FILMING_DATE_3 = "date46"
COLUMN_FILMING_DATE_4 = "date__1"
COLUMN_EDITING_TIME_1 = "date_range"
COLUMN_EDITING_TIME_2 = "date_range3"
COLUMN_EDITING_TIME_3 = "dup__of_editing_time_2__1"

# Time Calculation Rules
FILMING_HOURS_PER_DAY = 4
EDITING_HOURS_PER_DAY = 8
MAX_HOURS_PER_DAY = 8

# PowerPoint Styling (RGB colors)
HEADER_COLOR = (0, 112, 192)
ALT_ROW_COLOR = (217, 225, 242)
WHITE_COLOR = (255, 255, 255)
