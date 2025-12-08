################################
### Section 1: Initial Setup ###
################################

import pandas as pd
import numpy as np

from gurobipy import GRB, Model


FILE_PATH = "CoursePreferences.xlsx"

all_sheets = pd.read_excel(FILE_PATH, sheet_name=None)

# Data frames Courses, loads, and times
courses_df = all_sheets.get("Courses")
loads_df = all_sheets.get("Loads")
times_df = all_sheets.get("Times")

###############################
### Section 2: Attr Tables  ###
###############################

# Get the number of Prof, look at A and then full length
num_prof = courses_df.shape[1] - courses_df.columns.get_loc('A')

# Function converting integer to abc enum
def _index_to_col(num: int) -> str:
    """Convert idx to abc and excel style loop"""
    label = []
    while n > 0:
        n, r = divmod(n - 1, 26)
        label.append(chr(ord("A") + r))
    return "".join(reversed(label))

# Using conversion create labels for prof, output df
def df_with_letter_index(n: int) -> pd.DataFrame:
    """Create df with col loop"""
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
prof_attr['course_idx'] = prof_attr['course_idx'].astype(int)

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

#################################
### Section 3: Vars for Optim ###
#################################

### Pre-Determined
# a_i prof max credit loads
a_var = prof_attr.groupby("Prof")["max_credit"].mean().to_dict()

# b_j class total credits
b_var = courses_attr.set_index("Number")["Credits"].to_dict()

# c_{i,j} prof class elig
c_var = prof_attr.set_index(["Prof", "course_idx"])["course_bin"].to_dict()

# d_{i,k} prof time elig
d_var = prof_attr.set_index(["Prof", "time_idx"])["time_bin"].to_dict()

### Set Model

# Create the model
m = Model('course_sched') 

# Set parameters
m.setParam('OutputFlag', True)

### Decision Vars

# x_{i,j} prof to class
x_var = {
(i, j): m.addVar(name=f"x_{i}_{j}", lb=0)
for i in np.unique(np.array(prof_attr["Prof"])) for j in np.unique(np.array(prof_attr["course_idx"]))
}

# y_{j,k} class to time
y_var = {
(j, k): m.addVar(name=f"y_{j}_{}", lb=0)
for j in np.unique(np.array(prof_attr["course_idx"])) for k in np.unique(np.array(prof_attr["time_idx"]))
}

# z_{i,k} prof to time
z_var = {
(i, k): m.addVar(name=f"z_{i}_{k}", lb=0)
for i in np.unique(np.array(prof_attr["Prof"])) for k in np.unique(np.array(prof_attr["time_idx"]))
}

### Modelling

# Update model to integrate new variables
m.update()

# Constraint 1 ...
# Constraint 2 ...
# Constraint 3 ...
# Constraint 1 ...
# Constraint 4 ...
# Constraint 5 ...
# Constraint 6 ...
# Constraint 7 ...

# Update model to integrate new variables
m.update()








