################################
### Section 1: Initial Setup ###
################################

import pandas as pd
import numpy as np

import gurobipy as gp
from utils import df_with_letter_index, print_results


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
prof_attr['course_bin'] = prof_attr['course_bin'].astype(str).str.strip().replace('', 0)
prof_attr['course_bin'] = prof_attr['course_bin'].astype(int)

# Clean times attr table
times_attr['Times_Grp'] = times_attr['Times'].factorize()[0]
times_attr['Times_Grp_Day'] = (times_attr.index // 4) * 2 + (times_attr.index % 4 == 3).astype(int)

# Clean courses attr table
courses_attr.loc[courses_attr['Grad/Ugrad'] == 'Ugrad', 'Credits'] *= 3
courses_attr['Number'] = courses_attr['Number'].astype(int)
courses_attr['Number_Group'] = np.where(
    (courses_attr['Number'] >= 500) & (courses_attr['Number'] <= 699),
    6,
    ((courses_attr['Number'] - 1) // 100)
)
courses_attr['Labs'] = (courses_attr['Labs/Discussion Sections'] > 0).astype(int)

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

# d_{i,k} prof time elig
e_var = courses_attr.set_index("Number")["Labs"].to_dict()


### Set Model

# Create the model
m = gp.Model('course_sched') 

# Set parameters
m.setParam('OutputFlag', True)

### Decision Vars

# Idx arrays
idx_prof = np.unique(np.array(prof_attr["Prof"]))         # Set of Professors (i)
idx_course = np.unique(np.array(prof_attr["course_idx"])) # Set of Courses (j)
idx_time = np.unique(np.array(prof_attr["time_idx"]))     # Set of Time Slots (k)
course_groups = courses_attr.groupby('Number_Group')['Number'].apply(list).to_dict()
idx_group = course_groups.keys()                          # Set of Course Groups (g)
day_time_groups = times_attr.groupby('Times_Grp_Day')['index'].apply(list).to_dict()
idx_day_time = day_time_groups.keys()                     # Set of Day-Time Group IDs (t)
prime_indices = times_attr.iloc[4:16]['index'].tolist()   # Find index of prime class scheduling times
class_grp_701 = courses_attr[courses_attr['Number'] == 701]['Number_Group'].iloc[0] # Auto find 700 level classes
tues_thurs_indices = times_attr[times_attr['Days'] == "T/TH"]['index'].tolist() # Find index of tuesday/thursday classes

# x_{i,j} prof to class
x_var = {
(i, j): m.addVar(name=f"x_{i}_{j}", vtype=gp.GRB.BINARY)
for i in idx_prof for j in idx_course
}

# y_{j,k} class to time
y_var = {
(j, k): m.addVar(name=f"y_{j}_{k}", vtype=gp.GRB.BINARY)
for j in idx_course for k in idx_time
}

# z_{i,k} prof to time
z_var = {
(i, k): m.addVar(name=f"z_{i}_{k}", vtype=gp.GRB.BINARY)
for i in idx_prof for k in idx_time
}

# l_{j,k} lab to time
l_var = {
(i, j, k): m.addVar(name=f"l_{i}_{j}_{k}", vtype=gp.GRB.BINARY)
for i in idx_prof for j in idx_course for k in idx_time
}

### Modelling

# Update model to integrate new variables
m.update()

# Objective Func
m.setObjective(
    (gp.quicksum(x_var[i, j] * b_var[j] for i in idx_prof for j in idx_course)),
    gp.GRB.MAXIMIZE
)

# Constraint 1 ...
m.addConstrs(
    (gp.quicksum(x_var[i, j] for i in idx_prof) == 1
     for j in idx_course),
    name= f'one_prof'
)

# Constraint 2 ...
m.addConstrs(
    (gp.quicksum(x_var[i, j]*b_var[j] for j in idx_course) <= a_var[i]
     for i in idx_prof),
    name= f'prof_max'
)

# Constraint 3 ...
m.addConstrs(
    (x_var[i, j] <= c_var[i,j]
     for j in idx_course for i in idx_prof),
    name= f'prop_course'
)

# Constraint 4 ...
m.addConstrs(
    (gp.quicksum(y_var[j, k] for k in idx_time) == 1
     for j in idx_course),
    name= f'one_class'
)

# Constraint 5 ...
m.addConstrs(
    (gp.quicksum(z_var[i, k] for k in idx_time) <= -(-a_var[i]//3)
     for i in idx_prof),
    name= f'limit_prof'
)

# Constraint 6 ...
m.addConstrs(
    (z_var[i, k] <= d_var[i,k]
     for i in idx_prof for k in idx_time),
    name= f'proper_prof_time'
)

# Constraint 7 ...
m.addConstrs(
    (x_var[i, j] + y_var[j, k] - 1 <= z_var[i, k]
     for i in idx_prof for j in idx_course for k in idx_time),
    name= f'link_prof_class'
)

# Constraint 8 ...
m.addConstrs(
    (gp.quicksum(x_var[i, j] for i in idx_prof) == gp.quicksum(y_var[j, k] for k in idx_time)
     for j in idx_course),
    name= f'one_to_one'
)

# Constraint 9 ...
m.addConstrs(
    (gp.quicksum(l_var[i, j, k] for i in idx_prof for k in idx_time) == gp.quicksum(y_var[j, k]*e_var[j] for k in idx_time)
     for j in idx_course),
    name= f'lab_exists'
)

# Constraint 10 ...
m.addConstrs(
    (((1/3)*(x_var[i, j] + (1 - y_var[j, k]) + d_var[i, k])) >= l_var[i, j, k]
     for i in idx_prof for j in idx_course for k in idx_time),
    name= f'prof_lab_time'
)

# Constraint 11 ...
m.addConstrs(
    (gp.quicksum(y_var[j, k] for j in course_groups[g] for k in day_time_groups[t]) 
     + gp.quicksum(l_var[i, j, k] for i in idx_prof for j in course_groups[g] for k in day_time_groups[t]) <= 1
     for g in idx_group for t in idx_day_time),
    name= f'group_conflict'
)

# Constraint 12 ...
m.addConstr(
    y_var[701, 8] + y_var[701,10] == 1,
    name= f'grad_research_preffered'
)

# Constraint 13 ...
m.addConstrs(
    (gp.quicksum(y_var[j, k] for j in course_groups[g] for k in prime_indices if j != class_grp_701) * 2 <= 
     gp.quicksum(y_var[j, k] for j in course_groups[g] for k in idx_time if j != class_grp_701)
     for g in idx_group),
    name= f'max_50_percent_prime'
)

# Constraint 14 ...
m.addConstrs(
    (gp.quicksum(y_var[j, k] for j in course_groups[g] for k in tues_thurs_indices if j != class_grp_701) * 2 <= 
     gp.quicksum(y_var[j, k] for j in course_groups[g] for k in idx_time if j != class_grp_701)
     for g in idx_group),
    name= f'max_50_percent_T/Th'
)

# Update model to integrate new variables
m.update()

# Optimize the model
m.optimize()

# Print the result
status_codes = {
1: 'LOADED',
2: 'OPTIMAL',
3: 'INFEASIBLE',
4: 'INF_OR_UNBD',
5: 'UNBOUNDED'
}
status = m.status

#################################
### Section 4: Output Results ###
#################################

print_results(m, x_var, y_var, l_var, courses_attr, times_attr)