Attribute VB_Name = "NavigationModule"
Public Const student_data_sheet = "Student_Data"
Public Const advisor_data_sheet = "Advisor_Data"
Public Const course_conflict_data_sheet = "Course_Conflict_Data"
Public Const dashboard_sheet = "Dashboard"
Public Const add_students_sheet = "Add_Students"
Public Const student_matching_sheet = "Student_Matching"
Public Const advisor_schedule_sheet = "Advisor_Schedule"
Public Const general_stats_sheet = "General_Stats"
Public Const section_stats_sheet = "Section_Stats"

Sub GoToStudentData()
    'Subroutine that powers navigation buttons
    Sheets(student_data_sheet).Activate
    Sheets(student_data_sheet).Range("$A$1").Select
End Sub

Sub GoToAdvisorData()
    'Subroutine that powers navigation buttons
    Sheets(advisor_data_sheet).Activate
    Sheets(advisor_data_sheet).Range("$A$1").Select
End Sub

Sub GoToCourseConflictData()
    'Subroutine that powers navigation buttons
    Sheets(course_conflict_data_sheet).Activate
    Sheets(course_conflict_data_sheet).Range("$A$1").Select
End Sub

Sub GoToDashboard()
    'Subroutine that powers navigation buttons
    Sheets(dashboard_sheet).Activate
    Sheets(dashboard_sheet).Range("$A$1").Select
End Sub

Sub GoToAddStudents()
    'Subroutine that powers navigation buttons
    Sheets(add_students_sheet).Activate
    Sheets(add_students_sheet).Range("$A$1").Select
End Sub

Sub GoToStudentMatching()
    'Subroutine that powers navigation buttons
    Sheets(student_matching_sheet).Activate
    Sheets(student_matching_sheet).Range("$A$1").Select
End Sub

Sub GoToAdvisorSchedule()
    'Subroutine that powers navigation buttons
    Sheets(advisor_schedule_sheet).Activate
    Sheets(advisor_schedule_sheet).Range("$A$1").Select
End Sub

Sub GoToGeneralStats()
    'Subroutine that powers navigation buttons
    Sheets(general_stats_sheet).Activate
    Sheets(general_stats_sheet).Range("$A$1").Select
End Sub

Sub GoToSectionStats()
    'Subroutine that powers navigation buttons
    Sheets(section_stats_sheet).Activate
    Sheets(section_stats_sheet).Range("$A$1").Select
End Sub

