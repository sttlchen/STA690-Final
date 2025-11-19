import pandas as pd

FILE_PATH = "CoursePreferences.xlsx"

all_sheets = pd.read_excel(FILE_PATH, sheet_name=None)

# Show sheet names
print("Sheets found:", list(all_sheets.keys()))

courses_df = all_sheets.get("Courses")
loads_df = all_sheets.get("Loads")
times_df = all_sheets.get("Times")
conflicts_df = all_sheets.get("Conflicts")

print(courses_df.head())
print(loads_df.head())
print(times_df.head())
print(conflicts_df.head())