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

# Maximal course load for professor

prof_attr = prof_attr.merge(loads_df, on='Prof', how='left')
prof_attr = prof_attr.rename(columns={'NumCourses': 'max_credit'})
prof_attr['max_credit'] = prof_attr['max_credit'] * 3

# Melt the times and the courses matrix

times_df = times_df[:-2] # Issue with x=no in xlsx
times_df = times_df.reset_index()

# Create storage for later views on time

times_storage = times_df[['index','Times','Days']]
times_df = times_df.drop(columns=['Times', 'Days']).melt(id_vars='index', var_name='Prof', value_name='Value')
times_df = times_df.fillna(1).replace('x', 0)

# Join in times to the prof attr table

prof_attr = prof_attr.merge(times_df, on='Prof', how='left').rename(columns={'index': 'time_slot', 'Value': 'prof_time'})


print(prof_attr.head())
print(loads_df.head())
print(times_df.head())


