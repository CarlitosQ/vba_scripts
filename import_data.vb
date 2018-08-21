' From Martin Kresse @HHI

Public filename As String
Public intLeft As Integer, intRight As Integer

Sub FileSelect_Import_txt()


    DateiName1 = ActiveWorkbook.Name
    bolInit = True
    Dim makeGraphs As Boolean

    'speed improvment
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False

    ReadFiles = Application.GetOpenFilename("txt Files (*.txt),*.txt", , , , True)
    ' If ReadFiles = "Falsch" Then Exit Sub
    ' MsgBox "This selection consists of " & " files"

    'show userForm to interactivly change the sheetname
    filename = ReadFiles(1)
    UserForm1.Label2.Caption = filename
    UserForm1.Show

    For Each CurrentReadFile In ReadFiles
        ' copy Master-Sheet and select new Sheet
        Dim ws1 As Worksheet
        Set ws1 = ActiveWorkbook.Worksheets("Master")
        ws1.Copy After:=Sheets(Sheets.Count)
        Worksheets(Sheets.Count).Select
        ' Rename the Sheet
        ActiveSheet.[A1].Value = CurrentReadFile
        ActiveSheet.[B1].Value = renameSheet()

        On Error Resume Next
        ActiveSheet.Name = renameSheet()
        NoName:             If Err.Number = 1004 Then ActiveSheet.Name = renameSheet() & " (2)"
        If ActiveSheet.Name = ActNm Then GoTo NoName
        
        ' open and insert data
        Workbooks.OpenText filename:=CurrentReadFile, _
                           Origin:=xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
                           xlNone, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
                           Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1))
        
        DateiName2 = ActiveWorkbook.Name
        ActiveSheet.UsedRange.Select
        Selection.Copy
        
        Windows(DateiName1).Activate
        Worksheets(Sheets.Count).Select
        Range("A3").Select
        Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Range("A1").Select
        

        
        Application.DisplayAlerts = False
        Workbooks(DateiName2).Close , False
        Application.DisplayAlerts = True

    

    
    Next CurrentReadFile

    'speed improvment reverse
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True

End Sub

Function renameSheet() As String
    '
    ' renameSheets Makro
    Range("A1").Select
    inhalt = Left(ActiveCell.Formula, Len(ActiveCell.Formula) - intRight) 'cut the last letters off
    inhalt = Right(inhalt, intLeft)
    renameSheet = inhalt
End Function

Sub SortWksByCell()
    'Update 20141127
    Dim WorkRng As Range
    Dim WorkAddress As String
    On Error Resume Next
    xTitleId = "KutoolsforExcel"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range (Single)", xTitleId, WorkRng.Address, Type:=8)
    WorkAddress = WorkRng.Address
    Application.ScreenUpdating = False
    For i = 1 To Application.Worksheets.Count
        For j = i To Application.Worksheets.Count
            If VBA.UCase(Application.Worksheets(j).Range(WorkAddress)) < VBA.UCase(Application.Worksheets(i).Range(WorkAddress)) Then
                Application.Worksheets(j).Move Before:=Application.Worksheets(i)
            End If
        Next
    Next
    Application.ScreenUpdating = True
End Sub


'************************************************************************************
' Sub FileSelect_Import_txt_two_files_per_sheet()

'     DateiName1 = ActiveWorkbook.Name


'     bolInit = True
'     Dim makeGraphs As Boolean



'     ReadFiles = Application.GetOpenFilename("txt Files (*.txt),*.txt", , , , True)
'     ' If ReadFiles = "Falsch" Then Exit Sub
'     ' MsgBox "This selection consists of " & " files"

'     'show userFOrm to interactivly change the sheetname
'     filename = ReadFiles(1)
'     UserForm1.Label2.Caption = filename
'     UserForm1.Show

'     Dim intDoppler As Integer
'     intDoppler = 1

'     For Each CurrentReadFile In ReadFiles

'         If intDoppler Mod 2 = 1 Then

'             ' copy Master-Sheet and select new Sheet
'             Dim ws1 As Worksheet
'             Set ws1 = ActiveWorkbook.Worksheets("Master")
'             ws1.Copy After:=Sheets(Sheets.Count)
'             Worksheets(Sheets.Count).Select
'             ' Rename the Sheet
'             ActiveSheet.[A1].Value = CurrentReadFile
'             ActiveSheet.Range("A1").Select
'             MyName = Left(ActiveCell.Formula, Len(ActiveCell.Formula) - intRight) 'cut the last letters off
'             MyName = Right(inhalt, intLeft)
'             'ActiveSheet.Name = MyName
'             On Error Resume Next
'             NoName:             If Err.Number = 1004 Then ActiveSheet.Name = renameSheet() & " (2)"
'             If ActiveSheet.Name = ActNm Then GoTo NoName
        
'             ' open and insert data
'             Application.ScreenUpdating = False
'             Workbooks.OpenText filename:=CurrentReadFile, _
'                                Origin:=xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
'                                xlNone, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
'                                Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1))
        
'             DateiName2 = ActiveWorkbook.Name
'             ActiveSheet.UsedRange.Select
'             Selection.Copy
        
'             Windows(DateiName1).Activate
'             Worksheets(Sheets.Count).Select
'             Range("A3").Select
'             Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
'             Application.DisplayAlerts = False
'             Workbooks(DateiName2).Close , False
'             Application.DisplayAlerts = True
'             Application.ScreenUpdating = True
    
'         Else
    
'             Application.ScreenUpdating = False
'             ActiveSheet.[F1].Value = CurrentReadFile
'             Workbooks.OpenText filename:=CurrentReadFile, _
'                                Origin:=xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
'                                xlNone, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
'                                Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1))
        
'             DateiName2 = ActiveWorkbook.Name
'             ActiveSheet.UsedRange.Select
'             Selection.Copy
        
'             Windows(DateiName1).Activate
'             Worksheets(Sheets.Count).Select
'             Range("F3").Select
'             Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'             Range("A1").Select
        
'             Application.DisplayAlerts = False
'             Workbooks(DateiName2).Close , False
'             Application.DisplayAlerts = True
'             Application.ScreenUpdating = True
    
'         End If
'         intDoppler = intDoppler + 1
    
'     Next CurrentReadFile

' End Sub
