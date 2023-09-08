Attribute VB_Name = "mainMenuSubs"
Option Explicit
' ==== CP212 Windows Application Programming ===============+
' Name: Andrew Greco
' Student ID: 210422740
' Date: August 1, 2023
' Program title: Assignment 5
' Description: Student Marking Application
'===========================================================+

'Global file address and connection
Global name As String
Global db As New ADODB.Connection

Sub showMenu1()

    name = Application.GetOpenFilename
    With db
        .ConnectionString = "Data Source=" & name
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    mainMenu.Show

End Sub

'Callback for showMenuButton onAction
Sub showMenu(control As IRibbonControl)

    name = Application.GetOpenFilename
    With db
        .ConnectionString = "Data Source=" & name
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    mainMenu.Show
    
End Sub


Sub importData()
    Dim recordSet As New ADODB.recordSet
    Dim SQLQuery As String
    Dim dataSheet As Worksheet
    
    ' Clear and add a new worksheet named "Students" to store student data
    worksheetAddClear "Students"
    Set dataSheet = ThisWorkbook.Worksheets("Students")
    dataSheet.range("A1:C1").ColumnWidth = 10
    dataSheet.Cells(1, 1) = "First Name"
    dataSheet.Cells(1, 2) = "Last Name"
    dataSheet.Cells(1, 3) = "Student ID"
    
    ' populate the worksheet
    SQLQuery = "SELECT * FROM students"
    recordSet.Open SQLQuery, db
    dataSheet.range("A2").CopyFromRecordset recordSet
    recordSet.Close
    Set recordSet = Nothing
    
    ' Clear and add a new worksheet named "Grades" to store grade data
    worksheetAddClear "Grades"
    Set dataSheet = ThisWorkbook.Worksheets("Grades")
    dataSheet.range("A1:I1").ColumnWidth = 10
    dataSheet.Cells(1, 1) = "ID"
    dataSheet.Cells(1, 2) = "Student ID"
    dataSheet.Cells(1, 3) = "Course"
    dataSheet.Cells(1, 4) = "Assignment 1"
    dataSheet.Cells(1, 5) = "Assignment 2"
    dataSheet.Cells(1, 6) = "Assignment 3"
    dataSheet.Cells(1, 7) = "Assignment 4"
    dataSheet.Cells(1, 8) = "Midterm"
    dataSheet.Cells(1, 9) = "Exam"
    
    ' Retrieve grade data from the database
    SQLQuery = "SELECT * FROM grades"
    recordSet.Open SQLQuery, db
    dataSheet.range("A2").CopyFromRecordset recordSet
    recordSet.Close
    Set recordSet = Nothing
    
    ' Clear and add a new worksheet named "Courses" to store course data
    worksheetAddClear "Courses"
    Set dataSheet = ThisWorkbook.Worksheets("Courses")
    dataSheet.range("A1:C1").ColumnWidth = 10
    dataSheet.Cells(1, 1) = "ID"
    dataSheet.Cells(1, 2) = "Course Code"
    dataSheet.Cells(1, 3) = "Course Name"
    
    ' Retrieve course data from the database
    SQLQuery = "SELECT * FROM courses"
    recordSet.Open SQLQuery, db
    dataSheet.range("A2").CopyFromRecordset recordSet
    recordSet.Close
    Set recordSet = Nothing
End Sub


Private Sub worksheetAddClear(sheetName As String)
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ThisWorkbook
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        wb.Worksheets(sheetName).Delete
        Application.DisplayAlerts = True
    End If
    
    Dim newSheet As Worksheet
    Set newSheet = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    newSheet.name = sheetName
End Sub



Sub courseEnrollment()
    Dim SQLQuery As String
    Dim recordSet As New ADODB.recordSet
    Dim reportSheet As Worksheet
    
    ' Define the SQL queries based on selected courses
    If enrollCourse.as101.Value = True Then
        SQLQuery = "SELECT students.FirstName, students.LastName, grades.studentID FROM grades INNER JOIN students ON students.studentID = grades.studentID WHERE (grades.course='AS101')"
    ElseIf enrollCourse.cp102.Value = True Then
        SQLQuery = "SELECT students.FirstName, students.LastName, grades.studentID FROM grades INNER JOIN students ON students.studentID = grades.studentID WHERE (grades.course='CP102')"
    ElseIf enrollCourse.cp104.Value = True Then
        SQLQuery = "SELECT students.FirstName, students.LastName, grades.studentID FROM grades INNER JOIN students ON students.studentID = grades.studentID WHERE (grades.course='CP104')"
    ElseIf enrollCourse.cp212.Value = True Then
        SQLQuery = "SELECT students.FirstName, students.LastName, grades.studentID FROM grades INNER JOIN students ON students.studentID = grades.studentID WHERE (grades.course='CP212')"
    ElseIf enrollCourse.cp411.Value = True Then
        SQLQuery = "SELECT students.FirstName, students.LastName, grades.studentID FROM grades INNER JOIN students ON students.studentID = grades.studentID WHERE (grades.course='CP411')"
    ElseIf enrollCourse.pc120.Value = True Then
        SQLQuery = "SELECT students.FirstName, students.LastName, grades.studentID FROM grades INNER JOIN students ON students.studentID = grades.studentID WHERE (grades.course='PC120')"
    ElseIf enrollCourse.pc131.Value = True Then
        SQLQuery = "SELECT students.FirstName, students.LastName, grades.studentID FROM grades INNER JOIN students ON students.studentID = grades.studentID WHERE (grades.course='PC131')"
    ElseIf enrollCourse.pc141.Value = True Then
        SQLQuery = "SELECT students.FirstName, students.LastName, grades.studentID FROM grades INNER JOIN students ON students.studentID = grades.studentID WHERE (grades.course='PC141')"
    End If
    
    ' Clear and add new worksheet "Course Enrollment Report"
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ThisWorkbook
    On Error Resume Next
    Set ws = wb.Worksheets("CourseEnrollmentReport")
    On Error GoTo 0
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    
    Set reportSheet = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    reportSheet.name = "CourseEnrollmentReport"
    
    With reportSheet
        .Columns("A:C").ColumnWidth = 12
        .Cells(1, 1) = "First Name"
        .Cells(1, 2) = "Last Name"
        .Cells(1, 3) = "Student ID"
    End With
    
    ' Populating the report worksheet with the data
    recordSet.Open SQLQuery, db
    With reportSheet
        .range("A2").CopyFromRecordset recordSet
    End With
    recordSet.Close
    Set recordSet = Nothing
    
End Sub


Sub displayAverage()
   Dim recordSet As New ADODB.recordSet
    Dim SQLA1, SQLA2, SQLA3, SQLA4, SQLMidterm, SQLFinal As String
    Dim reportSheet As Worksheet
    Dim courseName As String

    ' Determine the selected course and corresponding SQL queries
    If classAverage.as101.Value = True Then
        courseName = "AS101 Astronomy I"
        SQLA1 = "SELECT AVG(A1) FROM grades WHERE (grades.course = 'AS101')"
        SQLA2 = "SELECT AVG(A2) FROM grades WHERE (grades.course = 'AS101')"
        SQLA3 = "SELECT AVG(A3) FROM grades WHERE (grades.course = 'AS101')"
        SQLA4 = "SELECT AVG(A4) FROM grades WHERE (grades.course = 'AS101')"
        SQLMidterm = "SELECT AVG(MidTerm) FROM grades WHERE (grades.course = 'AS101')"
        SQLFinal = "SELECT AVG(Exam) FROM grades WHERE (grades.course = 'AS101')"
        
    ElseIf classAverage.cp102.Value = True Then
        courseName = "CP102 Information Processing"
        SQLA1 = "SELECT AVG(A1) FROM grades WHERE course = 'CP102'"
        SQLA2 = "SELECT AVG(A2) FROM grades WHERE course = 'CP102'"
        SQLA3 = "SELECT AVG(A3) FROM grades WHERE course = 'CP102'"
        SQLA4 = "SELECT AVG(A4) FROM grades WHERE course = 'CP102'"
        SQLMidterm = "SELECT AVG(MidTerm) FROM grades WHERE course = 'CP102'"
        SQLFinal = "SELECT AVG(Exam) FROM grades WHERE course = 'CP102'"
        
    ElseIf classAverage.cp104.Value = True Then
        courseName = "CP104 Introduction to Programming"
        SQLA1 = "SELECT AVG(A1) FROM grades WHERE course = 'CP104'"
        SQLA2 = "SELECT AVG(A2) FROM grades WHERE course = 'CP104'"
        SQLA3 = "SELECT AVG(A3) FROM grades WHERE course = 'CP104'"
        SQLA4 = "SELECT AVG(A4) FROM grades WHERE course = 'CP104'"
        SQLMidterm = "SELECT AVG(MidTerm) FROM grades WHERE course = 'CP104'"
        SQLFinal = "SELECT AVG(Exam) FROM grades WHERE course = 'CP104'"
        
    ElseIf classAverage.cp212.Value = True Then
        courseName = "CP212 Windows Application Programming"
        SQLA1 = "SELECT AVG(A1) FROM grades WHERE course = 'CP212'"
        SQLA2 = "SELECT AVG(A2) FROM grades WHERE course = 'CP212'"
        SQLA3 = "SELECT AVG(A3) FROM grades WHERE course = 'CP212'"
        SQLA4 = "SELECT AVG(A4) FROM grades WHERE course = 'CP212'"
        SQLMidterm = "SELECT AVG(MidTerm) FROM grades WHERE course = 'CP212'"
        SQLFinal = "SELECT AVG(Exam) FROM grades WHERE course = 'CP212'"
        
    ElseIf classAverage.cp411.Value = True Then
        courseName = "CP411 Computer Graphics"
        SQLA1 = "SELECT AVG(A1) FROM grades WHERE course = 'CP411'"
        SQLA2 = "SELECT AVG(A2) FROM grades WHERE course = 'CP411'"
        SQLA3 = "SELECT AVG(A3) FROM grades WHERE course = 'CP411'"
        SQLA4 = "SELECT AVG(A4) FROM grades WHERE course = 'CP411'"
        SQLMidterm = "SELECT AVG(MidTerm) FROM grades WHERE course = 'CP411'"
        SQLFinal = "SELECT AVG(Exam) FROM grades WHERE course = 'CP411'"
        
    ElseIf classAverage.pc120.Value = True Then
        courseName = "PC120 Digital Electronics"
        SQLA1 = "SELECT AVG(A1) FROM grades WHERE course = 'PC120'"
        SQLA2 = "SELECT AVG(A2) FROM grades WHERE course = 'PC120'"
        SQLA3 = "SELECT AVG(A3) FROM grades WHERE course = 'PC120'"
        SQLA4 = "SELECT AVG(A4) FROM grades WHERE course = 'PC120'"
        SQLMidterm = "SELECT AVG(MidTerm) FROM grades WHERE course = 'PC120'"
        SQLFinal = "SELECT AVG(Exam) FROM grades WHERE course = 'PC120'"
        
    ElseIf classAverage.pc131.Value = True Then
        courseName = "PC131 Mechanics"
        SQLA1 = "SELECT AVG(A1) FROM grades WHERE course = 'PC131'"
        SQLA2 = "SELECT AVG(A2) FROM grades WHERE course = 'PC131'"
        SQLA3 = "SELECT AVG(A3) FROM grades WHERE course = 'PC131'"
        SQLA4 = "SELECT AVG(A4) FROM grades WHERE course = 'PC131'"
        SQLMidterm = "SELECT AVG(MidTerm) FROM grades WHERE course = 'PC131'"
        SQLFinal = "SELECT AVG(Exam) FROM grades WHERE course = 'PC131'"
        
    ElseIf classAverage.pc141.Value = True Then
        courseName = "PC141 Mechanics for Life Sciences"
        SQLA1 = "SELECT AVG(A1) FROM grades WHERE course = 'PC141'"
        SQLA2 = "SELECT AVG(A2) FROM grades WHERE course = 'PC141'"
        SQLA3 = "SELECT AVG(A3) FROM grades WHERE course = 'PC141'"
        SQLA4 = "SELECT AVG(A4) FROM grades WHERE course = 'PC141'"
        SQLMidterm = "SELECT AVG(MidTerm) FROM grades WHERE course = 'PC141'"
        SQLFinal = "SELECT AVG(Exam) FROM grades WHERE course = 'PC141'"
    End If
    
    ' Clear and add new worksheet "Course Average"
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ThisWorkbook
    On Error Resume Next
    Set ws = wb.Worksheets("CourseAverage")
    On Error GoTo 0
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    
    Set reportSheet = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    reportSheet.name = "CourseAverage"
    
    With reportSheet
        .Columns("A:G").ColumnWidth = 10
        .Cells(1, 1) = courseName
        .Cells(2, 2) = "A1"
        .Cells(2, 3) = "A2"
        .Cells(2, 4) = "A3"
        .Cells(2, 5) = "A4"
        .Cells(2, 6) = "Midterm"
        .Cells(2, 7) = "Final"
    End With
    
    ' Populating the report worksheet
    With recordSet
        .Open SQLA1, db
        reportSheet.range("B3").CopyFromRecordset recordSet
        .Close
        .Open SQLA2, db
        reportSheet.range("C3").CopyFromRecordset recordSet
        .Close
        .Open SQLA3, db
        reportSheet.range("D3").CopyFromRecordset recordSet
        .Close
        .Open SQLA4, db
        reportSheet.range("E3").CopyFromRecordset recordSet
        .Close
        .Open SQLMidterm, db
        reportSheet.range("F3").CopyFromRecordset recordSet
        .Close
        .Open SQLFinal, db
        reportSheet.range("G3").CopyFromRecordset recordSet
        .Close
    End With
    Set recordSet = Nothing

End Sub


