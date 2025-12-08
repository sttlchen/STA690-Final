import pandas as pd
import gurobipy as gp

# --- Helper Functions (From Section 2) ---

def _index_to_col(num: int) -> str:
    """Convert index number to Excel-style column label (A, B, C, AA, AB, etc.)."""
    label = []
    while num > 0:
        num, r = divmod(num - 1, 26)
        label.append(chr(ord("A") + r))
    return "".join(reversed(label))

def df_with_letter_index(n: int) -> pd.DataFrame:
    """Create a DataFrame with column labels representing professor indices (A, B, C...)."""
    labels = [_index_to_col(i) for i in range(1, n + 1)]
    return pd.DataFrame({'Prof': labels})

def safe_df_creation(data_list, columns):
    """Creates a DataFrame, ensuring column structure even if the list is empty."""
    if not data_list:
        return pd.DataFrame(columns=columns)
    return pd.DataFrame(data_list)


# --- Output Function (Replaces Section 4) ---

def print_results(m, x_var, y_var, l_var, courses_attr, times_attr):
    """
    Extracts, merges, and prints the course, time, and lab assignments
    from the optimized Gurobi model, and exports results to XLSX.
    """
    status_codes = {
        1: 'LOADED',
        2: 'OPTIMAL',
        3: 'INFEASIBLE',
        4: 'INF_OR_UNBD',
        5: 'UNBOUNDED'
    }

    if m.status == gp.GRB.OPTIMAL:
        
        # --- Data Extraction ---

        # 1. Extract Professor to Course Assignments (x_var)
        prof_course_assignments = []
        for (i, j), var in x_var.items():
            if var.X > 0.5:
                prof_course_assignments.append({'Prof': i, 'Course_Number': j})
        
        prof_course_df = safe_df_creation(prof_course_assignments, ['Prof', 'Course_Number'])

        prof_course_output = prof_course_df.merge(
            courses_attr[['Number', 'Name', 'Credits', 'Grad/Ugrad']], 
            left_on='Course_Number', 
            right_on='Number', 
            how='left'
        ).drop(columns=['Number'], errors='ignore') 

        # 2. Extract Course to Time Slot Assignments (y_var)
        course_time_assignments = []
        for (j, k), var in y_var.items():
            if var.X > 0.5:
                course_time_assignments.append({'Course_Number': j, 'Time_Index': k})
        
        course_time_df = safe_df_creation(course_time_assignments, ['Course_Number', 'Time_Index'])

        course_time_output = course_time_df.merge(
            times_attr[['index', 'Times', 'Days']],
            left_on='Time_Index',
            right_on='index',
            how='left'
        ).drop(columns=['index'], errors='ignore').merge(
            courses_attr[['Number', 'Name']],
            left_on='Course_Number',
            right_on='Number',
            how='left'
        ).drop(columns=['Number'], errors='ignore')

        # 3. Extract Lab Assignments (l_var)
        lab_assignments = []
        for (i, j, k), var in l_var.items():
            if var.X > 0.5:
                lab_assignments.append({'Prof': i, 'Course_Number': j, 'Time_Index': k})

        lab_df = safe_df_creation(lab_assignments, ['Prof', 'Course_Number', 'Time_Index'])
        
        lab_output = pd.DataFrame()
        if not lab_df.empty:
            lab_output = lab_df.merge(
                courses_attr[['Number', 'Name']],
                left_on='Course_Number',
                right_on='Number',
                how='left'
            ).drop(columns=['Number'], errors='ignore').merge(
                times_attr[['index', 'Times', 'Days']],
                left_on='Time_Index',
                right_on='index',
                how='left'
            ).drop(columns=['index'], errors='ignore')

        # 4. Combined Schedule (Prof, Course, Time)
        combined_schedule = pd.merge(
            prof_course_output, 
            course_time_output.drop(columns=['Name'], errors='ignore'),
            on='Course_Number', 
            how='inner'
        )
        
        # --- Terminal Output ---
        
        print("\n--- Professor-Course Assignments (x_var) ---")
        print(prof_course_output.sort_values(by=['Prof', 'Course_Number']))
        print("-" * 50)

        print("\n--- Course-Time Assignments (y_var) ---")
        print(course_time_output.sort_values(by=['Course_Number', 'Times']))
        print("-" * 50)

        if not lab_output.empty:
            print("\n--- Lab Assignments (l_var) ---")
            print(lab_output.sort_values(by=['Prof', 'Course_Number', 'Times']))
            print("-" * 50)
        else:
            print("\n--- No Lab Assignments Found ---")
            print("-" * 50)
        
        print("\n--- Final Course Schedule ---")
        final_cols = ['Prof', 'Name', 'Times', 'Days', 'Credits', 'Grad/Ugrad']
        present_cols = [col for col in final_cols if col in combined_schedule.columns]
        print(combined_schedule[present_cols].sort_values(by=['Prof', 'Times']))
        print("-" * 50)

        print(f"\nOptimal Objective Value (Total Credits Scheduled): {m.ObjVal}")

        # --- File Output (XLSX) ---
        
        output_filename = 'course_schedule_results.xlsx'
        print(f"\nWriting results to {output_filename}...")
        
        # Removed engine='xlsxwriter'
        with pd.ExcelWriter(output_filename) as writer: 
            prof_course_output.to_excel(writer, sheet_name='Prof_Course_Assignments', index=False)
            course_time_output.to_excel(writer, sheet_name='Course_Time_Assignments', index=False)
            combined_schedule.to_excel(writer, sheet_name='Combined_Schedule', index=False)
            if not lab_output.empty:
                lab_output.to_excel(writer, sheet_name='Lab_Assignments', index=False)
        print("Export complete.")

    else:
        print(f"\nOptimization ended with status: {status_codes.get(m.status, 'UNKNOWN')}")
        if m.status == gp.GRB.INFEASIBLE:
            print("Model is infeasible. Consider computing IIS (m.computeIIS()) to debug constraints.")