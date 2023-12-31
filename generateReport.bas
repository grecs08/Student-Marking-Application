Attribute VB_Name = "generateReport"
Option Explicit

Sub generateReport()

' Defintions
    Dim rs As New ADODB.recordSet
    Dim SQLA1, SQLA2, SQLA3, SQLA4, SQLMidterm, SQLFinal As String
    Dim sheet As Worksheet
    Dim className As String
    Dim ch As ChartObject
    Dim rn As range
    Dim SQLStatements As Variant
    Dim ColumnLabels As Variant
    
' Data selection
    If report.as101.Value = True Then
        className = "AS101"
        SQLA1 = "SELECT AVG(A1) FROM grades WHERE (grades.course = 'AS101')"
        SQLA2 = "SELECT AVG(grades.A2) FROM grades WHERE (grades.course = 'AS101')"
        SQLA3 = "SELECT AVG(grades.A3) FROM grades WHERE (grades.course = 'AS101')"
        SQLA4 = "SELECT AVG(grades.A4) FROM grades WHERE (grades.course = 'AS101')"
        SQLMidterm = "SELECT AVG(grades.MidTerm) FROM grades WHERE (grades.course = 'AS101')"
        SQLFinal = "SELECT AVG(grades.Exam) FROM grades WHERE (grades.course = 'AS101')"
        
    ElseIf report.cp102.Value = True Then
        className = "CP102"
        SQLA1 = "SELECT AVG(A1) FROM grades WHERE course = 'CP102'"
        SQLA2 = "SELECT AVG(A2) FROM grades WHERE course = 'CP102'"
        SQLA3 = "SELECT AVG(A3) FROM grades WHERE course = 'CP102'"
        SQLA4 = "SELECT AVG(A4) FROM grades WHERE course = 'CP102'"
        SQLMidterm = "SELECT AVG(MidTerm) FROM grades WHERE course = 'CP102'"
        SQLFinal = "SELECT AVG(Exam) FROM grades WHERE course = 'CP102'"
        
    ElseIf report.cp104.Value = True Then
        className = "CP104"
        SQLA1 = "SELECT AVG(A1) FROM grades WHERE course = 'CP104'"
        SQLA2 = "SELECT AVG(A2) FROM grades WHERE course = 'CP104'"
        SQLA3 = "SELECT AVG(A3) FROM grades WHERE course = 'CP104'"
        SQLA4 = "SELECT AVG(A4) FROM grades WHERE course = 'CP104'"
        SQLMidterm = "SELECT AVG(MidTerm) FROM grades WHERE course = 'CP104'"
        SQLFinal = "SELECT AVG(Exam) FROM grades WHERE course = 'CP104'"
        
    ElseIf report.cp212.Value = True Then
        className = "CP212"
        SQLA1 = "SELECT AVG(A1) FROM grades WHERE course = 'CP212'"
        SQLA2 = "SELECT AVG(A2) FROM grades WHERE course = 'CP212'"
        SQLA3 = "SELECT AVG(A3) FROM grades WHERE course = 'CP212'"
        SQLA4 = "SELECT AVG(A4) FROM grades WHERE course = 'CP212'"
        SQLMidterm = "SELECT AVG(MidTerm) FROM grades WHERE course = 'CP212'"
        SQLFinal = "SELECT AVG(Exam) FROM grades WHERE course = 'CP212'"
        
    ElseIf report.cp411.Value = True Then
        className = "CP411"
        SQLA1 = "SELECT AVG(A1) FROM grades WHERE course = 'CP411'"
        SQLA2 = "SELECT AVG(A2) FROM grades WHERE course = 'CP411'"
        SQLA3 = "SELECT AVG(A3) FROM grades WHERE course = 'CP411'"
        SQLA4 = "SELECT AVG(A4) FROM grades WHERE course = 'CP411'"
        SQLMidterm = "SELECT AVG(MidTerm) FROM grades WHERE course = 'CP411'"
        SQLFinal = "SELECT AVG(Exam) FROM grades WHERE course = 'CP411'"
        
    ElseIf report.pc120.Value = True Then
        className = "PC120"
        SQLA1 = "SELECT AVG(A1) FROM grades WHERE course = 'PC120'"
        SQLA2 = "SELECT AVG(A2) FROM grades WHERE course = 'PC120'"
        SQLA3 = "SELECT AVG(A3) FROM grades WHERE course = 'PC120'"
        SQLA4 = "SELECT AVG(A4) FROM grades WHERE course = 'PC120'"
        SQLMidterm = "SELECT AVG(MidTerm) FROM grades WHERE course = 'PC120'"
        SQLFinal = "SELECT AVG(Exam) FROM grades WHERE course = 'PC120'"
        
    ElseIf report.pc131.Value = True Then
        className = "PC131"
        SQLA1 = "SELECT AVG(A1) FROM grades WHERE course = 'PC131'"
        SQLA2 = "SELECT AVG(A2) FROM grades WHERE course = 'PC131'"
        SQLA3 = "SELECT AVG(A3) FROM grades WHERE course = 'PC131'"
        SQLA4 = "SELECT AVG(A4) FROM grades WHERE course = 'PC131'"
        SQLMidterm = "SELECT AVG(MidTerm) FROM grades WHERE course = 'PC131'"
        SQLFinal = "SELECT AVG(Exam) FROM grades WHERE course = 'PC131'"
        
    ElseIf report.pc141.Value = True Then
        className = "PC141"
        SQLA1 = "SELECT AVG(A1) FROM grades WHERE course = 'PC141'"
        SQLA2 = "SELECT AVG(A2) FROM grades WHERE course = 'PC141'"
        SQLA3 = "SELECT AVG(A3) FROM grades WHERE course = 'PC141'"
        SQLA4 = "SELECT AVG(A4) FROM grades WHERE course = 'PC141'"
        SQLMidterm = "SELECT AVG(MidTerm) FROM grades WHERE course = 'PC141'"
        SQLFinal = "SELECT AVG(Exam) FROM grades WHERE course = 'PC141'"
    End If
    
  ' Define the SQL statements and column labels in arrays
    SQLStatements = Array(SQLA1, SQLA2, SQLA3, SQLA4, SQLMidterm, SQLFinal)
    ColumnLabels = Array("A1", "A2", "A3", "A4", "Midterm", "Exam")

    ' Clear and add a new worksheet with the given name
    clearOneWorksheet (className)
    Call worksheetSubs.newWorksheet(className)

    ' Set the worksheet and apply some formatting
    Set sheet = ThisWorkbook.Worksheets(className)
    With sheet.range("A1")
        .Style = "Title"
        .ColumnWidth = 30
        .Font.Size = 20
    End With

    sheet.Cells(1, 1) = className & " Report"

    ' Populate data using a loop
    Dim i As Integer
    Dim dataRange As range

    With rs
        For i = 0 To UBound(SQLStatements)
            .Open SQLStatements(i), db
            Set dataRange = sheet.range("C" & (i + 2))
            dataRange.CopyFromRecordset rs
            .Close
        Next i
    End With

    ' Set the column labels
    For i = 0 To UBound(ColumnLabels)
        sheet.Cells(i + 2, 2) = ColumnLabels(i)
    Next i

    ' Set the data range for the chart
    Set rn = sheet.range("B2:C7")

    ' Plotting chart
    Set ch = sheet.ChartObjects.Add(Left:=325, Width:=400, Top:=46, Height:=250)
    With ch.Chart
        .SetSourceData rn
        .ChartType = xlBarClustered
        .HasTitle = True
        .ChartTitle.Text = "Evaluation Averages"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Evaluations"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Grades (%)"
        .Legend.Delete
    End With
    
End Sub

Sub createWordDoc()

' Error Handling
    On Error Resume Next
    
' Defintions
    Dim app As Word.Application
    Dim doc As Word.Document
    Dim ws As Worksheet
    
' Creating word application and document
    Set app = CreateObject("Word.Application")
    app.Visible = True
    Set doc = app.Documents.Add
    app.ActiveDocument.SaveAs2 "Grade_Report_Average"


' Document Formatting
    With app.Selection
        .Font.Size = 24
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .TypeText ("Grade Report Average")
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .TypeParagraph
        .Font.Size = 14
        .TypeText ("This is a Microsoft Word report compiled from average grades of various courses." & vbNewLine)
        .Font.Size = 18
        .TypeParagraph
    End With
    
    For Each ws In Worksheets
        If IsCourseWorksheet(ws) Then
            app.Selection.TypeText (vbNewLine)
            app.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            app.Selection.Font.Bold = True
            app.Selection.Font.Size = 16
            app.Selection.TypeText ("Course: " & ws.name & vbNewLine)
            app.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            app.Selection.Font.Size = 12
            
            ' Copy data from the worksheet and create sepration
            ws.UsedRange.Copy
            doc.ActiveWindow.Selection.PasteExcelTable False, False, False
            app.Selection.TypeText (vbNewLine & "--------------------------------------------------------------------------------------------" & vbNewLine)
        End If
    Next ws

    With doc.Content
        .ParagraphFormat.SpaceAfter = 0
        .Font.Size = 11
    End With

    On Error GoTo 0
    
End Sub

Function IsCourseWorksheet(ws As Worksheet) As Boolean
    Dim courseNames As Variant
    courseNames = Array("AS101", "CP102", "CP104", "CP212", "CP411", "PC120", "PC131", "PC141")
    IsCourseWorksheet = Not IsError(Application.Match(ws.name, courseNames, 0))
End Function

rades WHERE course = 'CP104'"
        SQLFinal = "SELECT AVG(Exam) FROM grades WHERE course = 'CP104'"
        
    ElseIf classAverage.cp212.Value = True Then
        courseName = "CP212 Windows Application Programming"
        SQLA1 = "SELECT AVG(A1) 