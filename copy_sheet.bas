Sub copy_sheet()
    '
    ' copy_sheet Macro
    '

    '
    Dim ws As Worksheet
    Dim NewName As String
    Set ws = ThisWorkbook.Worksheets("Master")
    NewName = InputBox("New Name:")
    ws.Copy After:=ThisWorkbook.Sheets(Sheets.Count)
    ActiveSheet.Name = NewName

End Sub


'*********************************************************************
Sub copy_sheet()
    '
    ' copy_sheet Macro
    '

    '
    Dim ws As Worksheet
    Dim NewName As String
    Dim CopyFrom As String
    CopyFrom = ActiveSheet.Name
    Set ws = ThisWorkbook.Worksheets(CopyFrom)
    ws.Copy After:=ThisWorkbook.Sheets(Sheets.Count)

End Sub

