Attribute VB_Name = "MainUpdateCode"
Public Const error_print_data = "ErrorPrint.txt"
Public Const solution_data = "SolutionOutput.txt"
Public Const student_output_data = "StudentOutput.txt"
Public Const advisor_schedule_data = "AdvisorScheduleOutput.txt"
Public Const student_matching_sheet = "Student_Matching"
Public Const advisor_schedule_sheet = "Advisor_Schedule"
Public Const general_stats_sheet = "General_Stats"
Public Const section_stats_sheet = "Section_Stats"
Public Const solution_output_sheet = "Solution_Output"
Public Const dashboard_sheet = "Dashboard"
Public Const student_data_sheet = "Student_Data"
Public Const course_conflict_sheet = "Course_Conflict_Data"
Public Const advisor_data_sheet = "Advisor_Data"

    


Function dept_code(long_dept) As String
    'Function that converts the long department name into the code that is used for the majors
    If long_dept = "Applied and Engineering Physics" Or long_dept = "Engineering Physics" Or long_dept = "EP" Then
        dept_code = "EP"
    ElseIf long_dept = "Biological and Environmental Engineering" Or long_dept = "BE" Then
        dept_code = "BE"
    ElseIf long_dept = "Biomedical Engineering" Or long_dept = "BME" Then
        dept_code = "BME"
    ElseIf long_dept = "Chemical and Biomolecular Engineering" Or long_dept = "CHEME" Then
        dept_code = "CHEME"
    ElseIf long_dept = "Civil Engineering" Or long_dept = "CE" Then
        dept_code = "CE"
    ElseIf long_dept = "Computer Science" Or long_dept = "CS" Then
        dept_code = "CS"
    ElseIf long_dept = "Earth and Atmospheric Sciences" Or long_dept = "SES" Then
        dept_code = "SES"
    ElseIf long_dept = "Electrical and Computer Engineering" Or long_dept = "ECE" Then
        dept_code = "ECE"
    ElseIf long_dept = "Environmental Engineering" Or long_dept = "EnvirE" Then
        dept_code = "EnvirE"
    ElseIf long_dept = "Information Science Systems and Technology" Or long_dept = "ISST" Then
        dept_code = "ISST"
    ElseIf long_dept = "Materials Science and Engineering" Or long_dept = "MSE" Then
        dept_code = "MSE"
    ElseIf long_dept = "Mechanical and Aerospace Engineering" Or long_dept = "ME" Then
        dept_code = "ME"
    ElseIf long_dept = "Operations Research and Information Engineering" Or long_dept = "OR" Then
        dept_code = "OR"
    
    Return

End Function



Sub Run_Python()
    'Call the start.cmd file which changes the file path and called the Python file
    workbook_path = ActiveWorkbook.Path
    RetVal = Shell(workbook_path & "\start.cmd")

End Sub

Sub check_data()
    'Routine that checks data to see if there are
    'issues with the data
    
    'Count the amount the data in the Student_Data table
    'Number of headings
    num_headings = Range("Student_Headings").Count
    'Number of data cells
    total_cells = Range("Student_Headings").CurrentRegion.Count - num_headings - 1
    
    'Declare error_offset for all errors
    error_offset = 1
    
    If num_headings = total_cells Then
    'There's no data
        Error = MsgBox("Error: There is no data in the Student_Data range on sheet Student_Data." _
            & "Please enter data before running.", vbCritical, "Abort")
    End If
    
    'Count the amount of data in the Advisor_Data table
    'Number of headings
    num_headings = WorksheetFunction.CountA(Range("Advisor_Headings"))
    'Number of data cells
    total_cells = WorksheetFunction.CountA(Range("Advisor_Data"))
    
    If num_headings = total_cells Then
    'There's no data
        Error = MsgBox("Error: There is no data in the Advisor_Data range on sheet Advisor_Data." _
            & "Please enter data before running.", vbCritical, "Abort")
    End If
    
    'Count the amount of data in the Course_Conflict_Data table
    'Number of headings
    num_headings = WorksheetFunction.CountA(Range("Course_Conflict_Headings"))
    'Number of data cells
    total_cells = WorksheetFunction.CountA(Range("Course_Conflict_Data"))
    
    If num_headings = total_cells Then
    'There's no data
        Error = MsgBox("Error: There is no data in the Course_Conflict_Data range on " _
            & "sheet Course_Conflict_Data.  Please enter data before running.", vbCritical, "Abort")
    End If

    'Convert the times in the "Start Time" and "End Time" columns on
    'the course conflict data sheet to the correc printing format
    'Find the columns
    start_col = WorksheetFunction.Match("Start Time", Range("Course_Conflict_Headings"), 0)
    end_col = WorksheetFunction.Match("End Time", Range("Course_Conflict_Headings"), 0)
    start_cell = Range("Course_Conflict_Headings").Cells(1, start_col).Address
    end_cell = Range("Course_Conflict_Headings").Cells(1, end_col).Address
    num_courses = Range(Range("Course_Conflict_Headings"), Range("Course_Conflict_Headings").End(xlDown)).Rows.Count - 1
    
    'Iterate through on the course conflict sheet and overwrite the start times
    For i = 1 To num_courses
        'Start time
        Sheets(course_conflict_sheet).Range(start_cell).Offset(i, 0).NumberFormat = "General"
        Sheets(course_conflict_sheet).Range(start_cell).Offset(i, 0).Value = "'" & _
            WorksheetFunction.Text(Sheets(course_conflict_sheet).Range(start_cell).Offset(i, 0).Value, "hh:mm:ss AM/PM")
        
        'End Time
        Sheets(course_conflict_sheet).Range(end_cell).Offset(i, 0).NumberFormat = "General"
        Sheets(course_conflict_sheet).Range(end_cell).Offset(i, 0).Value = "'" & _
            WorksheetFunction.Text(Sheets(course_conflict_sheet).Range(end_cell).Offset(i, 0).Value, "hh:mm:ss AM/PM")
        
    Next

    'Need to check that all students have at least 1 point assigned to them
    'Need to take away two rows since there's the heading row
    'and also the warning row
    num_students = Sheets(student_data_sheet).Range("Student_Headings").CurrentRegion.Rows.Count - 2
    For i = 1 To num_students
        total_points = WorksheetFunction.Sum(Sheets(student_data_sheet).Range("Student_Headings").Offset(i, 1))
        If total_points = 0 Then
            student_id = Sheets(student_data_sheet).Range("student_headings").Offset(i, 0)(0)
            'MsgBox ("Error: Student " & student_id & _
            '    " needs to have at least one point assigned to a major.  Please correct and try again.")
            'Want to print to the Error box
            Sheets(dashboard_sheet).Range("Error_Printing").Offset(error_offset, 0).Value = "Error: Student " & student_id & _
                " needs to have at least one point assigned to a major.  "
            'Increment the error_offset
            error_offset = error_offset + 1
            Exit Sub
        End If
    Next

End Sub

Sub create_csv()
    'Get the activeworkbook name
    workbook_name = ActiveWorkbook.FullName
    workbook_path = ActiveWorkbook.Path

    'Eliminate screen flicker
    Application.ScreenUpdating = False
    
    'Subroutine to create and export .csv files
    'that the script/python need to run processs
    'into .dat files for AMPL/Gurobi
    
    'Make the full_student_data.csv file
    student_filename = "New_Full_Student_Data.csv"
    'Copy the data into the New_Full_Student_Data sheet
    Range("Student_Data").Copy
    Sheets("New_Full_Student_Data").Visible = True
    Sheets("New_Full_Student_Data").Range("A1").PasteSpecial (xlPasteValues)
    
    'Need to turn off alerts to just overwrite the file
    Application.DisplayAlerts = False
    Sheets("New_Full_Student_Data").SaveAs Filename:=workbook_path + ("\") + student_filename, FileFormat:=xlCSV
    Application.DisplayAlerts = True
    
    'Cleanup
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs workbook_name, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    Sheets("New_Full_Student_Data").Range("A1").CurrentRegion.ClearContents
    Sheets("New_Full_Student_Data").Visible = False
    
    'Make the Advisor_Preference_Data.csv file
    advisor_filename = "Advisor_Preference_Data.csv"
    'Copy the data into the Advisor_Preference_Data sheet
    Range(Range("Advisor_Headings"), Range("Advisor_Headings").End(xlDown)).Copy
    Sheets("Advisor_Preference_Data").Visible = True
    Sheets("Advisor_Preference_Data").Range("A1").PasteSpecial (xlPasteValues)
    
    'Need to append all the times
    advisor_header = Range(Sheets("Advisor_Preference_Data").Range("$A$1"), _
                Sheets("Advisor_Preference_Data").Range("$A$1").End(xlToRight)).Address
    mon_col = WorksheetFunction.Match("Monday Times", Sheets("Advisor_Preference_Data").Range(advisor_header), 0) - 1
    tues_col = WorksheetFunction.Match("Tuesday Times", Sheets("Advisor_Preference_Data").Range(advisor_header), 0) - 1
    wed_col = WorksheetFunction.Match("Wednesday Times", Sheets("Advisor_Preference_Data").Range(advisor_header), 0) - 1
    thur_col = WorksheetFunction.Match("Thursday Times", Sheets("Advisor_Preference_Data").Range(advisor_header), 0) - 1
    fri_col = WorksheetFunction.Match("Friday Times", Sheets("Advisor_Preference_Data").Range(advisor_header), 0) - 1
    
    num_rows = Range(Sheets("Advisor_Preference_Data").Range("$A$1").Offset(1, 0), _
                     Sheets("Advisor_Preference_Data").Range("$A$1").End(xlDown)).Rows.Count
    Sheets("Advisor_Preference_Data").Range("$A$1").End(xlToRight).Offset(0, 1).Value = "Advisor_Times"
    
    'write the formula for the first time
    For i = 1 To num_rows
        
        first_comma = ""
        second_comma = ""
        third_comma = ""
        fourth_comma = ""
        
        'get the times for each of the days of the week
        mon_value = Sheets("Advisor_Preference_Data").Range("$A$1").Offset(i, mon_col).Value
    
        tues_value = Sheets("Advisor_Preference_Data").Range("$A$1").Offset(i, tues_col).Value
    
        wed_value = Sheets("Advisor_Preference_Data").Range("$A$1").Offset(i, wed_col).Value
    
        thur_value = Sheets("Advisor_Preference_Data").Range("$A$1").Offset(i, thur_col).Value
        
        fri_value = Sheets("Advisor_Preference_Data").Range("$A$1").Offset(i, fri_col).Value
        
        'Need to check if we need the commas
        If mon_value <> "" And (tues_value <> "" Or wed_value <> "" Or thur_value <> "" Or fri_value <> "") Then
            first_comma = ","
        End If
    
        If tues_value <> "" And (wed_value <> "" Or thur_value <> "" Or fri_value <> "") Then
            second_comma = ","
        End If
    
        If wed_value <> "" And (thur_value <> "" Or fri_value <> "") Then
            third_comma = ","
        End If
        
        If thur_value <> "" And fri_value <> "" Then
            fourth_comma = ","
        End If
    
    Sheets("Advisor_Preference_Data").Range("$A$1").End(xlToRight).Offset(i, 0).Value = _
        mon_value & first_comma & tues_value & second_comma & wed_value & _
            third_comma & thur_value & fourth_comma & fri_value
    Next
    
    
    'Need to turn off alerts to just overwrite the file
    Application.DisplayAlerts = False
    Sheets("Advisor_Preference_Data").SaveAs Filename:=workbook_path + ("\") + advisor_filename, FileFormat:=xlCSV
    Application.DisplayAlerts = True
    Sheets("Advisor_Preference_Data").Range("A1").CurrentRegion.ClearContents
    
    'Cleanup
    Sheets("Advisor_Preference_Data").Visible = False
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs workbook_name, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    
    'Make the Course_Conflict_Data.csv file
    course_conflict_filename = "Course_Conflict_Data_Sheet.csv"
    'Copy the data into the Course_Conflict_Data sheet
    Range("Course_Conflict_Data").Copy
    Sheets("Course_Conflict_Data_Sheet").Visible = True
    Sheets("Course_Conflict_Data_Sheet").Range("A1").PasteSpecial (xlPasteValues)
    
    'Need to turn off alerts to just overwrite the file
    Application.DisplayAlerts = False
    Sheets("Course_Conflict_Data_Sheet").SaveAs workbook_path + "\" + course_conflict_filename, FileFormat:=xlCSV
    Sheets("Course_Conflict_Data_Sheet").Range("A1").CurrentRegion.ClearContents
    Sheets("Course_Conflict_Data_Sheet").Visible = False
    Application.DisplayAlerts = True
    
    'Cleanup
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=workbook_name, FileFormat:=52
    Application.DisplayAlerts = True
    
    
    Application.ScreenUpdating = True

End Sub

Sub Run_Full_Algorithm()
'Run the full matching algorithm from start to finish

    'Run basic checks
    Call check_data

    'First have to create the csvs
    Call create_csv
    
    'Then call the algorithm
    Call Run_Python

End Sub

Sub Import_All_Data()
    'Need a separate subroutine that the user has to press
    'when the AMPL file is completed

    'Turn off screen updating
    Application.ScreenUpdating = False

    'Import the student match
    Call Import_Student_Matching
    
    'Import the advisor schedule
    Call Import_Advisor_Schedule
    
    'Import the Solution Output for pivots
    Call Import_Solution_Output
    
    'Import the errors from the dat file
    'Call Import_Error_Print
    
    'Refresh all pivots
    Sheets(general_stats_sheet).PivotTables("PivotTable1").ChangePivotCache ActiveWorkbook. _
        PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Sheets(solution_output_sheet).Range("$A$1").CurrentRegion.Address _
        , Version:=6)
    Sheets(general_stats_sheet).PivotTables("PivotTable2").ChangePivotCache ("PivotTable1")
    Sheets(section_stats_sheet).PivotTables("PivotTable3").ChangePivotCache ( _
        general_stats_sheet & "!PivotTable2")
    Sheets(section_stats_sheet).PivotTables("PivotTable4").ChangePivotCache ( _
        general_stats_sheet & "!PivotTable2")
    Sheets(section_stats_sheet).PivotTables("PivotTable5").ChangePivotCache ( _
        general_stats_sheet & "!PivotTable2")
        
    'Go back to the dashboard
    Call GoToDashboard

    'Turn on screen updating
    Application.ScreenUpdating = True

End Sub


Sub Import_Student_Matching()
    'Subroutine to import the Student matching
    'data into the corresponding
    'sheet for easier use
    'also used for adding the extra students
    
    'Turn off screen updating
    Application.ScreenUpdating = False
    
    'File should be in the same directory as the workbook from other subroutines
    workbook_path = ActiveWorkbook.Path
    
    'Delete if there is anything
    Application.DisplayAlerts = False
    Sheets(student_matching_sheet).Range("Student_Matching_Start").CurrentRegion.Clear
    Application.DisplayAlerts = True
    
    'Activate sheet--Excel rules, apparently
    Sheets(student_matching_sheet).Activate
    
    'Import
    With Sheets(student_matching_sheet).QueryTables.Add(Connection:= _
        "TEXT;" & workbook_path & "\" & student_output_data _
        , Destination:=Range("Student_Matching_Start"))
        .Name = "StudentOutput"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    'Delete the connection
    ActiveWorkbook.Connections("StudentOutput").Delete
    
    'Turn on screen updating
    Application.ScreenUpdating = True

End Sub

Sub Import_Error_Print()
    'Subroutine to import the presolve errors from the advisor matching
    'or advisor times
    
    'Turn off screen updating
    Application.ScreenUpdating = False
    
    'File should be in the same directory as the workbook from other subroutines
    workbook_path = ActiveWorkbook.Path
    
    'Clear the current errors if any
    Range("Error_Printing").CurrentRegion.ClearContents
    
    'Activate the sheet in case
    Sheets(dashboard_sheet).Activate
    
    'Import the Error Printing file
    With Sheets(dashboard_sheet).QueryTables.Add(Connection:= _
        "TEXT;" & workbook_path & "\" & error_print_data _
        , Destination:=Range("Error_Printing"))
        .Name = "ErrorPrint"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    'Delete the connection
    ActiveWorkbook.Connections("ErrorPrint").Delete
    
    'Turn on screen updating
    Application.ScreenUpdating = True

End Sub

Sub Import_Advisor_Schedule()
    'Subroutine to import the Advisor schedule
    'data into the corresponding sheet for easier use

    'Turn off screen updating
    Application.ScreenUpdating = False

    'File should be in the same directory as the workbook from other subroutines
    workbook_path = ActiveWorkbook.Path
    
    'Delete if there is anything
    Application.DisplayAlerts = False
    Sheets(advisor_schedule_sheet).Range("Advisor_Schedule_Start").CurrentRegion.Clear
    Application.DisplayAlerts = True

    'Activate sheet, just in case
    Sheets(advisor_schedule_sheet).Activate

    'Import data
    With Sheets(advisor_schedule_sheet).QueryTables.Add(Connection:= _
        "TEXT;" & workbook_path & "\" & advisor_schedule_data _
        , Destination:=Range("Advisor_Schedule_Start"))
        .Name = "AdvisorScheduleOutput"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1250
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    'Delete the connection
    ActiveWorkbook.Connections("AdvisorScheduleOutput").Delete

    'Turn on screen updating
    Application.ScreenUpdating = True

End Sub

Sub Import_Solution_Output()
    'Subroutine to import the full Solution Output
    'similar to how we do it for the template
    
    'Turn off screen updating
    Application.ScreenUpdating = False
    
    'Unhide sheet
    Sheets(solution_output_sheet).Visible = True
    
    'File should be in the same directory as the workbook from other subroutines
    workbook_path = ActiveWorkbook.Path
    
    'Clear the current data--turn off the alert just for this command
    Application.DisplayAlerts = False
    Sheets(solution_output_sheet).Range("$A$1").CurrentRegion.Clear
    Application.DisplayAlerts = True
    
    'Activate sheet, just in case
    Sheets(solution_output_sheet).Activate
    
    'Import
    With Sheets(solution_output_sheet).QueryTables.Add(Connection:= _
        "TEXT;" & workbook_path & "\" & solution_data _
        , Destination:=Sheets(solution_output_sheet).Range("$A$1"))
        .Name = "SolutionOutput"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With

    'Delete the connection
    ActiveWorkbook.Connections("SolutionOutput").Delete
    
    'Hide sheet
    Sheets(solution_output_sheet).Visible = False
    
    'Turn on screen updating
    Application.ScreenUpdating = True

End Sub


Sub Refresh_Pivot()
    'Subroutine to refresh all the pivots that
    'are based off the Solution output sheet
    
    'Turn off screen updating
    Application.ScreenUpdating = False
    
    Sheets(general_stats_sheet).PivotTables("PivotTable1").ChangePivotCache ActiveWorkbook. _
        PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Sheets(solution_output_sheet).Range("$A$1").CurrentRegion.Address _
        , Version:=6)
    Sheets(general_stats_sheet).PivotTables("PivotTable2").ChangePivotCache ("PivotTable1")
    Sheets(section_stats_sheet).PivotTables("PivotTable3").ChangePivotCache ( _
        general_stats_sheet & "!PivotTable1")
    Sheets(section_stats_sheet).PivotTables("PivotTable4").ChangePivotCache ( _
        general_stats_sheet & "!PivotTable1")
    Sheets(section_stats_sheet).PivotTables("PivotTable5").ChangePivotCache ( _
        general_stats_sheet & "!PivotTable1")
    
    'ActiveWorkbook.RefreshAll
    For Each PT In Sheets(general_stats_sheet).PivotTables
        PT.RefreshTable
    Next PT

    For Each PT In Sheets(section_stats_sheet).PivotTables
        PT.RefreshTable
    Next PT

    'Turn on screen updating
    Application.ScreenUpdating = True
End Sub

Sub Add_Student_to_Section()
    'Subroutine to add the extra students to sections that only have 14 students
    
    'Turn off screen updating
    Application.ScreenUpdating = False
    
    'Declare dictionary
    Set advisor_count_dictionary = New Scripting.Dictionary
    
    'First check if the students are already in the matching
    'Add_Student_Header is the header of the list of students
    'Take away 1 cell for the header
    num_new_students = Range("Add_Student_Header").CurrentRegion.Count - 1
    
    'List of advisors is on the advisor_schedule_sheet
    num_advisors = Range(Sheets(advisor_schedule_sheet).Range("$A$5"), _
        Sheets(advisor_schedule_sheet).Range("Advisor_Schedule_Start").End(xlDown)).Rows.Count - 1
    
    'Add all the keys to dictionary
    For i = 1 To num_advisors
        'Check if the key exists, if not then add it with value of 0
        advisor_key = Sheets(advisor_schedule_sheet).Range("Advisor_Schedule_Start").Offset(i, 0).Value
        
        If Not advisor_count_dictionary.Exists(advisor_key) Then
            advisor_count_dictionary.Add advisor_key, 0
        End If
        
    Next
    
    'Find the advisors who have 14 students--just iterate through the student matching
    num_students = Sheets(student_matching_sheet).Range("Student_Matching_Start").CurrentRegion.Rows.Count - 1
    
    'Iterate through students and count them per advisor; the advisor is in column B
    For i = 1 To num_students
        'Get the advisor name
        advisor_name = Sheets(student_matching_sheet).Range("Student_Matching_Start").Offset(i, 1).Value
        
        'Increment the key in the dictionary
        advisor_count_dictionary(advisor_name) = advisor_count_dictionary(advisor_name) + 1
    Next
    
    
    'Go through each of the students to see if there are duplicates
    For i = 1 To num_new_students
        count_occurrence = WorksheetFunction.CountIf(Range( _
            Sheets(student_matching_sheet).Range("Student_Matching_Start"), _
            Sheets(student_matching_sheet).Range("Student_Matching_Start").End(xlDown)), _
            "=" & Range("Add_Student_Header").Offset(i, 0))
        If count_occrrence <> 0 Then
            MsgBox ("Error: Student ID " & Range("Add_Student_Header").Offset(i, 0).Value & "already in matching. " _
                & "Please remove from the list of students to add and try again.")
            'End early
            Exit Sub
        'Else we need to find an advisor that has 14 students or as few students as possible
        'and add the student to that advisor on the sheet
        Else
            'Iterate through and find the min value and the associated advisor
            min_value = 100 'Really high upper bound for number of students per advisor
            min_advisor = ""
            For Each k In advisor_count_dictionary.Keys()
                '14 is the lower bound from the model as per advising office parameters
                If advisor_count_dictionary(k) = 14 Then
                    
                    'Assign
                    min_advisor = k
                    min_value = 14
                    Exit For
                    
                ElseIf advisor_count_dictionary(k) < min_value Then
                    min_value = advisor_count_dictionary(k)
                    min_advisor = k
                End If
            Next k
            
            'Went through all advisors and the min is not 14 so just add the student
            'to the min advisor that we already found
            'Add the student value
            Sheets(student_matching_sheet).Range("Student_Matching_Start").End(xlDown).Offset(1, 0).Value = _
                        Range("Add_Student_Header").Offset(i, 0).Value
            'Add the advisor
            Sheets(student_matching_sheet).Range("Student_Matching_Start").End(xlDown).Offset(0, 1).Value = min_advisor
            
            'Add the student to the advisor value
            advisor_count_dictionary(min_advisor) = advisor_count_dictionary(min_advisor) + 1
            
            'Add the student to the end of SolutionOutput
            Sheets(solution_output_sheet).Range("$A$1").End(xlDown).Offset(1, 0).Value = Range("Add_Student_Header").Offset(i, 0).Value
            
            'Add the advisor to the end of Solution Output
            Sheets(solution_output_sheet).Range("$A$1").End(xlDown).Offset(0, 1).Value = min_advisor
            
            'Refresh pivots that depend on SolutionOutput
            Call Refresh_Pivot
        End If
    
    Next
    
    'Hide the SolutionOutput sheet
    Sheets(solution_output_sheet).Visible = False
    
    'Turn on screen updating
    Application.ScreenUpdating = True
    
    'Go to the Student Matching sheet if not already there
    Call GoToStudentMatching
    
End Sub

Sub check_advisor_data()
    'Subroutine that checks if the advisor data is all feasible
    
    Call create_csv
    
    'Call the python code
    RetVal = Shell("python " & workbook_path & "Check_advisor_python.py")
    
    'Import the erros
    Call Import_Error_Print
    
    'Turn on screen updating
    Application.ScreenUpdating = True
End Sub
