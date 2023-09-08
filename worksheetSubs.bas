Attribute VB_Name = "worksheetSubs"
Sub worksheetClear()
    Dim ws As Worksheet
    For Each ws In Worksheets
        If (ws.name <> "Home") Then
            Application.DisplayAlerts = False
            Sheets(ws.name).Delete
            Application.DisplayAlerts = True
        End If
    Next ws
End Sub

'Callback for clearWorksheetsButton onAction
Sub worksheetsClear(control As IRibbonControl)
 Dim ws As Worksheet
    For Each ws In Worksheets
        If (ws.name <> "Home") Then
            Application.DisplayAlerts = False
            Sheets(ws.name).Delete
            Application.DisplayAlerts = True
        End If
    Next ws
End Sub

Sub newWorksheet(sheetName As String)
     With ThisWorkbook
        .Worksheets.Add(After:=.Sheets(.Sheets.Count)).name = sheetName
    End With
End Sub

Sub clearOneWorksheet(name As String)
    Dim ws As Worksheet
    For Each ws In Worksheets
        If (ws.name = name) Then
            Application.DisplayAlerts = False
            Sheets(ws.name).Delete
            Application.DisplayAlerts = True
        End If
    Next ws
End Sub

   �+���          ��(۰                          0              heming   Ҹ�                                 �+���                                  h  X  @  V           inkingManager                                 �+���                            �  �  X  ��������         �ɰ  anager                 ��        �-� Њ�߰  @}丰          @  0    Z  ��������    �          