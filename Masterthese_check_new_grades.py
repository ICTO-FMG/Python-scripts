from canvasapi import Canvas
import yaml
import pandas as pd
import datetime
import numpy as np
import smtplib
import win32com.client as win32

# Load the config file
with open('config.yaml') as data:
    canvas_config = yaml.safe_load(data)['canvas']

# Set the base URL for the Canvas API
API_URL = canvas_config['url']

# Set the access token for the Canvas API
API_KEY = canvas_config['accesstoken']

# Initialize a new Canvas object
canvas = Canvas(API_URL, API_KEY)

# Create lists of course codes that you want the script to look at
unique_codes_thesis = ["7204MTKP", "7204KPXXXX", "7204MBACXY", "7205RMTXY"] # hardcoded, old thesis pages that still need to be checked
standard_codes_thesis = ["MTXXXXB&C", "MTXXXXKNP", "MTXXXXBCiS", "MTXXXXKFP", "MTXXXXKPS1", "MTXXXXKPS2", "MTXXXXKLOPS1", "MTXXXXKLOPS2", "MTXXXXKLOP", "MTXXXXW&O", "MTXXXXHP&BC", 
                         "MTXXXXBDS", "MTXXXXSP", "MTXXXXRM"] # standardized course codes which we will replace the XXXX'es of to get the codes of courses created by the Python script Masterthesis_page_creation
standard_codes_internsh = ["7205RMIXY", "MIXXXXRM", "MIXXXXBDS", "MIXXXXB&C", "MIXXXXHP&BC", "MIXXXXW&O", "MIXXXXSP", "MIXXXXT&D"]                    

academic_year_filter = ("2122", "2223", "2324", "2425") # At the start of each new academic year, add that year

# Create a vector of course codes in which the XXXX's from the standard codes are replaced by the specified academic years
new_thesis_courses = unique_codes_thesis + [sub.replace('XXXX', academic_year_filter[0]) for sub in standard_codes_thesis] + [sub.replace('XXXX', academic_year_filter[1]) for sub in standard_codes_thesis] + [sub.replace('XXXX', academic_year_filter[2]) for sub in standard_codes_thesis] + [sub.replace('XXXX', academic_year_filter[3]) for sub in standard_codes_thesis]
new_internsh_courses = [sub.replace('XXXX', academic_year_filter[1]) for sub in standard_codes_internsh] + [sub.replace('XXXX', academic_year_filter[2]) for sub in standard_codes_internsh] + [sub.replace('XXXX', academic_year_filter[3]) for sub in standard_codes_internsh]

all_courses = canvas.get_account(123).get_courses(state = ['available']) # Get all published courses from the PSY subaccount (123)
all_courses_df = pd.DataFrame([vars(course) for course in all_courses]) # put paginated list in df

# Filter on some old thesis courses
thesis_courses_filtered = all_courses_df[(all_courses_df["name"].str.contains("thes|Thes")) & (all_courses_df["sis_course_id"].str.contains("2191|2201")) 
                            & (all_courses_df["course_code"].str.contains("7204|7205"))] # These vakken uit 19/20 en 20/21 die niet in unique_codes_thesis zitten maar waarvan er wel nog een paar gecheckt moeten worden

# Filter on the thesis and internship courses in new_thesis_courses and new_internsh_courses
thesis_courses_new = all_courses_df[(all_courses_df["course_code"].isin(new_thesis_courses))]
internsh_courses_new = all_courses_df[(all_courses_df["course_code"].isin(new_internsh_courses))]

# Concat thesis DataFrames
thesis_courses_filtered = pd.concat([thesis_courses_filtered, thesis_courses_new])
internsh_courses_filtered = internsh_courses_new

## THESIS
final_grades_thesis = pd.DataFrame() # initialize empty dataframe
success = 0 # initialize success parameter for final rowbinding. This is to prevent errors when there are no grades available.

for i in range(len(thesis_courses_filtered)): # for all thesis courses

  course = canvas.get_course(thesis_courses_filtered["id"].iloc[i]) # set the course that you want to look up the grades for
  students = course.get_users(enrollment_role='StudentEnrollment') # get the students in this course
  students_data = pd.DataFrame([vars(stud) for stud in students]) # put paginated list in df
  
  if students_data.empty: # check whether there are students in the course
    print('There are no students in course ',thesis_courses_filtered["id"].iloc[i])
  else:
    students_data = students_data[["id", "sis_user_id", "sortable_name"]].drop_duplicates() # select relevant student info
    assignments = course.get_assignments() # get assignment info
    assignments_df = pd.DataFrame([vars(ass) for ass in assignments]) # put paginated list in df
    
    # Filter assignments on relevant assignments
    grade_assignments_filtered = assignments_df[(assignments_df['name'].str.contains("6. Final grade")) | (assignments_df['name'].str.contains("6r. Resit:"))] # in these assignments the grades are entered
    thesis_names = assignments_df[(assignments_df['name'].str.contains("5. Submit")) | (assignments_df['name'].str.contains("5r. Resit:"))] # in these assignments the students submit the final version
   
   # Check if there's a 'Final check education desk' column, if not, create one
    if len(assignments_df[assignments_df['name'].str.contains("Final check")]) == 0:
      course.create_assignment(assignment = {"name": "Final check education desk", "grading_type": "pass_fail", "published": True})
      assignments_df = pd.DataFrame([vars(ass) for ass in course.get_assignments()]) # re-create the assignments DataFrame for it to contain the newly created assignment
      education_desk_check = assignments_df[assignments_df['name'].str.contains("Final check")] # get the information for the education desk check column
    else:
      education_desk_check = assignments_df[assignments_df['name'].str.contains("Final check")]
    
    # For each step 6 assignment, get the grades
    for k in range(len(grade_assignments_filtered)):
      assignment = course.get_assignment(grade_assignments_filtered["id"].iloc[k]) # set the assignment
      subm = assignment.get_submissions()
      grade_data = pd.DataFrame([vars(ass) for ass in subm]) # put paginated list in df
      
      if grade_data.empty:
        print('There are no grades available for course',thesis_courses_filtered["id"].iloc[i],'- assignment',grade_assignments_filtered["name"].iloc[k])
        if k == 0:
          students_data[['final_grade', 'graded_at', 'submit_status']] = None # create a DataFrame with None's
        else: # if grade data IS empty and k != 0
          students_data[['resit_grade', 'resit_graded_at', 'resit_submit_status']] = None # create a resit DataFrame with None's
      else: # if grade_data of the 6/6r assignment is not empty
        assignment = course.get_assignment(thesis_names["id"].iloc[k]) # set corresponding step 5/5r assignment
        subm = assignment.get_submissions()
        thesis_info = pd.DataFrame([vars(ass) for ass in subm]) # get submission info of 5/5r assignment
        thesis_info = thesis_info[["user_id", "workflow_state"]] # just keep the user_id and workflow_state
        grade_data = grade_data[["user_id", "grade", "graded_at"]] # from the submission info of 6/6r assignment, keep user_id, grade and date of grading
        grade_data = grade_data.merge(thesis_info, how='left', on='user_id') # combine the info
        if k == 0: # then you're looking at a 1st chance
          grade_data.columns = ["user_id", "final_grade", "graded_at", "submit_status"]
        else: # then you're looking at the resit
          grade_data.columns = ["user_id", "resit_grade", "resit_graded_at", "resit_submit_status"]
        students_data = students_data.merge(grade_data, left_on='id', right_on='user_id', how='left') # add grade information to students_data DataFrame
 
    ## Finalize data for one course by adding a course name, a course link and an education desk check column
    success = 1
    students_data['track'] = thesis_courses_filtered["name"].iloc[i] # add a 'track' column with the coursename
    students_data['link'] = "https://canvas.uva.nl/courses/"+str(thesis_courses_filtered["id"].iloc[i]) # add a 'link' column with the url to the Canvas course
    students_data = students_data.drop('user_id', axis=1, errors='ignore')
    
    assignment = course.get_assignment(education_desk_check["id"].iloc[0]) # set the education desk check assignment
    subm = assignment.get_submissions() # get the results
    check_data = pd.DataFrame([vars(ass) for ass in subm]) # put the results from the education desk check column in a DataFrame
    check_data = check_data[["user_id", "grade"]] # just keep the user_id and grade
    check_data.columns = ["user_id", "Balie_Check"] # rename grade to Balie_Check
    students_data = students_data.merge(check_data, left_on='id', right_on='user_id', how='left') # add Balie_Check info

  while success == 1:
   final_grades_thesis = pd.concat([final_grades_thesis, students_data]) # now concat students_data to final_grades_thesis and start the loop again
   success = 0

# Now clean up the final_grades_thesis dataframe
final_grades_thesis = final_grades_thesis.dropna(subset =['final_grade', 'resit_grade'], how = 'all') # drop all rows with no final grade AND no resit grade
final_grades_thesis = final_grades_thesis.drop('id', axis=1)
final_grades_thesis = final_grades_thesis.sort_values('graded_at', ascending=False)

final_grades_thesis['graded_at'] = pd.to_datetime(final_grades_thesis['graded_at'], format="%Y-%m-%dT%H:%M:%SZ")
final_grades_thesis['resit_graded_at'] = pd.to_datetime(final_grades_thesis['resit_graded_at'], format="%Y-%m-%dT%H:%M:%SZ")

filter_date = datetime.datetime(2021, 3, 1, 0, 0, 0)

final_grades_thesis = final_grades_thesis[(final_grades_thesis["graded_at"] > filter_date) | (final_grades_thesis["resit_graded_at"] > filter_date)]
final_grades_thesis = final_grades_thesis[final_grades_thesis['Balie_Check'].isna()] # only keep the rows where Balie_Check is NA, otherwhise the grade has already been entered to SIS
final_grades_thesis = final_grades_thesis.drop(["graded_at", "Balie_Check", "user_id", "resit_graded_at", "user_id_x", "user_id_y"], axis=1) # drop columns
final_grades_thesis.columns = ["Student ID", "Full Name", "Final_Grade", "Submit_Status","Resit_Grade","Resit_Submit_Status","Track", "Link"] # rename columns
final_grades_thesis['Submit_Status'] = final_grades_thesis['Submit_Status'].str.replace('graded','submitted')
final_grades_thesis['Resit_Submit_Status'] = final_grades_thesis['Resit_Submit_Status'].str.replace('graded','submitted')
final_grades_thesis = final_grades_thesis.reset_index(drop=True)
final_grades_thesis = final_grades_thesis.fillna(value=np.nan)
final_grades_thesis = final_grades_thesis.fillna('')

## INTERNSHIP
final_grades_internship = pd.DataFrame() # initialize empty dataframe
success = 0 # initialize success parameter for final rowbinding. This is to prevent errors when there are no grades available.

for i in range(len(internsh_courses_filtered)): # for all internship courses

  course = canvas.get_course(internsh_courses_filtered["id"].iloc[i]) # set the course that you want to look up the grades for
  students = course.get_users(enrollment_role='StudentEnrollment') # get the students in this course
  students_data = pd.DataFrame([vars(stud) for stud in students]) # put paginated list in df
  
  if students_data.empty: # check whether there are students in the course
    print('There are no students in course ',internsh_courses_filtered["id"].iloc[i])
  else:
    students_data = students_data[["id", "sis_user_id", "sortable_name"]].drop_duplicates() # select relevant student info
    assignments = course.get_assignments() # get assignment info
    assignments_df = pd.DataFrame([vars(ass) for ass in assignments]) # put paginated list in df
    assignments_df = assignments_df[assignments_df["id"]!= 423847] # probleem met bepaalde assignment van internship social psychology

    # Filter assignments on relevant assignments
    if "RM" in internsh_courses_filtered["course_code"].iloc[i]: # then it's a resmas course and there's an other column to be checked
      grade_assignments_filtered = assignments_df[(assignments_df['name'].str.contains("6. Final grade")) | (assignments_df['name'].str.contains("6r. Resit:"))]
    else:
      grade_assignments_filtered = assignments_df[assignments_df['name'].str.contains("Upload FINAL internship report")]
   
    education_desk_check = assignments_df[assignments_df['name'].str.contains("Final check")] # look up the 'Final check education desk' column
    
    # For each grade_assignments_filtered assignment, get the grades
    for k in range(len(grade_assignments_filtered)):
      assignment = course.get_assignment(grade_assignments_filtered["id"].iloc[k]) # set the assignment
      subm = assignment.get_submissions()
      grade_data = pd.DataFrame([vars(ass) for ass in subm]) # put paginated list in df
      
      if grade_data.empty or (len(grade_data.index)==grade_data.grade.isnull().sum()):
        print('There are no grades available for course',internsh_courses_filtered["id"].iloc[i],'- assignment',grade_assignments_filtered["name"].iloc[k])
        if k == 0:
          students_data[['final_grade', 'graded_at', 'submit_status']] = None # create a DataFrame with None's
        else: # if grade data IS empty and k != 0
          students_data[['resit_grade', 'resit_graded_at', 'resit_submit_status']] = None # create a resit DataFrame with None's
      else: # if grade_data is not empty
        assignment = course.get_assignment(grade_assignments_filtered["id"].iloc[k]) # get assignment
        subm = assignment.get_submissions()
        thesis_info = pd.DataFrame([vars(ass) for ass in subm]) # get submission info of assignment
        thesis_info = thesis_info[["user_id", "workflow_state"]] # just keep the user_id and workflow_state
        grade_data = grade_data[["user_id", "grade", "graded_at"]] # from the submission info, keep user_id, grade and date of grading
        grade_data = grade_data.merge(thesis_info, how='left', on='user_id') # combine the info
        if k == 0: # then you're looking at a 1st chance
          grade_data.columns = ["user_id", "final_grade", "graded_at", "submit_status"]
        else: # then you're looking at the resit
          grade_data.columns = ["user_id", "resit_grade", "resit_graded_at", "resit_submit_status"]
        students_data = students_data.merge(grade_data, left_on='id', right_on='user_id', how='left') # add grade information to students_data DataFrame
 
    ## Finalize data for one course by adding course name, course link and education desk check column
    success = 1
    students_data['track'] = internsh_courses_filtered["name"].iloc[i] # add a 'track' column with the coursename
    students_data['link'] = "https://canvas.uva.nl/courses/"+str(internsh_courses_filtered["id"].iloc[i]) # add a 'link' column with the url of the course
    students_data = students_data.drop('user_id', axis=1, errors='ignore')
    
    assignment = course.get_assignment(education_desk_check["id"].iloc[0]) # set the education desk check assignment
    subm = assignment.get_submissions() # get the results
    check_data = pd.DataFrame([vars(ass) for ass in subm]) # put the results in the education desk check column in a DataFrame
    check_data = check_data[["user_id", "grade"]] # just keep the user_id and grade
    check_data.columns = ["user_id", "Balie_Check"] # rename grade to Balie_Check
    students_data = students_data.merge(check_data, left_on='id', right_on='user_id', how='left') # add Balie_Check info

  while success == 1:
   final_grades_internship = pd.concat([final_grades_internship, students_data]) # now concat students_data to final_grades_thesis and start the loop again
   success = 0

# Now clean up the final_grades_internship dataframe
final_grades_internship = final_grades_internship.dropna(subset =['final_grade', 'resit_grade'], how = 'all') # drop all rows with no final grade AND no resit grade
final_grades_internship = final_grades_internship.drop('id', axis=1)
final_grades_internship = final_grades_internship.sort_values('graded_at', ascending=False)

final_grades_internship['graded_at'] = pd.to_datetime(final_grades_internship['graded_at'], format="%Y-%m-%dT%H:%M:%SZ")
final_grades_internship['resit_graded_at'] = pd.to_datetime(final_grades_internship['resit_graded_at'], format="%Y-%m-%dT%H:%M:%SZ")

filter_date = datetime.datetime(2021, 3, 1, 0, 0, 0)

final_grades_internship = final_grades_internship[(final_grades_internship["graded_at"] > filter_date) | (final_grades_internship["resit_graded_at"] > filter_date)]
final_grades_internship = final_grades_internship[final_grades_internship['Balie_Check'].isna()] # only keep the rows where Balie_Check is NA, otherwhise the grade has already been entered to SIS
final_grades_internship = final_grades_internship.drop(["graded_at", "Balie_Check", "user_id", "resit_graded_at", "user_id_x", "user_id_y"], axis=1) # drop columns
final_grades_internship.columns = ["Student ID", "Full Name", "Final_Grade", "Submit_Status","Resit_Grade","Resit_Submit_Status","Track", "Link"] # rename columns
final_grades_internship['Submit_Status'] = final_grades_internship['Submit_Status'].str.replace('graded','submitted')
final_grades_internship['Resit_Submit_Status'] = final_grades_internship['Resit_Submit_Status'].str.replace('graded','submitted')
final_grades_internship = final_grades_internship.reset_index(drop=True)
final_grades_internship = final_grades_internship.fillna(value=np.nan)
final_grades_internship = final_grades_internship.fillna('')

## SENDING EMAIL
if final_grades_thesis.empty and final_grades_internship.empty:
  print("There are no new grades to report.")
else: # send email
  Outlook = win32.Dispatch('outlook.application')
  
  df_to_html_thesis = final_grades_thesis.to_html()
  df_to_html_interns = final_grades_internship.to_html()
  
  if (not final_grades_thesis.empty) and (not final_grades_internship.empty):
    body = "<html><p>The following students have handed in their thesis/internship and are available for the check.</p><p> Thesis pages: </p>"+df_to_html_thesis+"<p> Internship pages: </p>"+df_to_html_interns+"<p>This is an automatically generated email.</p></html>"
  elif (final_grades_thesis.empty) and (not final_grades_internship.empty):
    body = "<html><p>The following students have handed in their thesis/internship and are available for the check.</p><p> Thesis pages: none </p>"+"<p> Internship pages: </p>"+df_to_html_interns+"<p>This is an automatically generated email.</p></html>"
  elif (not final_grades_thesis.empty) and (final_grades_internship.empty):
    body = "<html><p>The following students have handed in their thesis/internship and are available for the check.</p><p> Thesis pages: </p>"+df_to_html_thesis+"<p> Internship pages: none </p>"+"<p>This is an automatically generated email.</p></html>"

  Email = Outlook.CreateItem(0)
  Email.To = 'afstudeerpsychologie-fmg@uva.nl'
  Email.CC = 'icto-psy@uva.nl'
  Email.Subject = 'New Mastertheses/internships available for final check'
  Email.HTMLBody = body

  Email.Send()
