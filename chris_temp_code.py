### Initial Setup

import pandas as pd

FILE_PATH = "CoursePreferences.xlsx"

all_sheets = pd.read_excel(FILE_PATH, sheet_name=None)

# Show sheet names
print("Sheets found:", list(all_sheets.keys()))

# Data frames Courses, loads, and times
courses_df = all_sheets.get("Courses")
loads_df = all_sheets.get("Loads")
times_df = all_sheets.get("Times")

### Develop dataframe for the Professor attr

# Get the number of Prof, look at A and then full length
num_prof = courses_df.shape[1] - courses_df.columns.get_loc('A')

# Function converting integer to abc enum
def _index_to_col(num: int) -> str:
    s = ''
    num, rem = divmod(num - 1, 26)
    s = chr(ord('A') + rem) + s
    return s

# Using conversion create labels for prof, output df
def df_with_letter_index(n: int) -> pd.DataFrame:
    labels = [_index_to_col(i) for i in range(1, n + 1)]
    return pd.DataFrame({'Prof': labels})

# Create inital prof data.frame
prof_attr = df_with_letter_index(num_prof)

print(courses_df.head())
print(loads_df.head())
print(times_df.head())


