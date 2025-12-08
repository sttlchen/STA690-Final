### Initial Setup

import pandas as pd
import numpy as np

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
times_attr = times_df[['index','Times','Days']] # Create times attr
times_df = times_df.drop(columns=['Times', 'Days']).melt(id_vars='index', var_name='Prof', value_name='Value')
times_df = times_df.fillna(1).replace('x', 0) # x = no

courses_df = courses_df[:-2] #issue with xlsx
courses_attr = courses_df[['Number','Name','Grad/Ugrad','Credits','Labs/Discussion Sections','Total Enrollment']] # Create courses attr
courses_df = courses_df.drop(columns=['Name','Grad/Ugrad','Credits','Labs/Discussion Sections','Total Enrollment']).melt(id_vars='Number', var_name='Prof', value_name='Value')
courses_df = courses_df.fillna(0).replace('x', 1) # x = yes



# Join in times, courses to the prof attr table

prof_attr = prof_attr.merge(times_df, on='Prof', how='left').rename(columns={'index': 'time_idx', 'Value': 'time_bin'})
prof_attr = prof_attr.merge(courses_df, on='Prof', how='left').rename(columns={'Number': 'course_idx', 'Value': 'course_bin'})

# Clean times attr table

times_attr['Times_Grp'] = times_attr['Times'].factorize()[0]
times_attr_temp = times_attr.assign(Day=times_attr['Days'].str.split('/')).explode('Day')
times_attr['Times_Grp_Day'] = pd.factorize(
    times_attr_temp.groupby(['Times', 'Day']).ngroup()
    .groupby(times_attr_temp.index).min()
)[0]

# Clean courses attr table

courses_attr.loc[courses_attr['Grad/Ugrad'] == 'Ugrad', 'Credits'] *= 3
courses_attr['Number'] = courses_attr['Number'].astype(int)
courses_attr['Number_Group'] = np.where(
    courses_attr['Number'].isin([521, 523, 532, 581]),
    8,
    ((courses_attr['Number'] - 1) // 100)
)

### Set Decision and Pre-Determined Vars



