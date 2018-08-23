' For intendation: http://rubberduckvba.com/Indentation

Sub import_txt()
    '
    ' import_txt Macro
    '

    '
    Dim sPath As String
    Dim NewName As String
    Dim NewConnection As String
    Dim ws As Worksheet
    Dim CopyFrom As String
    
    ' Copy current sheet first
    CopyFrom = ActiveSheet.Name
    Set ws = ThisWorkbook.Worksheets(CopyFrom)
    ws.Copy After:=ThisWorkbook.Sheets(Sheets.Count)

    ' Select target file and import
    sPath = Application.GetOpenFilename()
    NewName = InputBox("New Sheet Name:")
    NewConnection = "TEXT;" & sPath
    
    With ActiveSheet
        .Cells.ClearContents
        .Name = NewName
    End With

    With ActiveSheet.QueryTables.Add(Connection:=NewConnection _
                                                  , Destination:=Range("$A$1"))
        ' .CommandType = 0
        .Name = NewName
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
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
End Sub

