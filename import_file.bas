Option Explicit

Sub ImportFile()
    Dim sPath As String
    Dim CurrentSheetName As String
    'Below we assume that the file, csvtest.csv,
    'is in the same folder as the workbook. If
    'you want something more flexible, you can
    'use Application.GetOpenFilename to get a
    'file open dialogue that returns the name
    'of the selected file.
    'On the page Fast text file import
    'I show how to do that - just replace the
    'file pattern "txt" with "csv".
    sPath = Application.GetOpenFilename()
    ' sPath = ThisWorkbook.Path & "\csvtest.csv"

    'Procedure call. Semicolon is defined as separator,
    'and data is to be inserted on "Sheet2".
    'Of course you could also read the separator
    'and sheet name from the worksheet or an input
    'box. There are several options.
    CurrentSheetName = ActiveSheet.Name
    ' copyDataFromCsvFileToSheet sPath, ";", "Sheet1"
    copyDataFromCsvFileToSheet sPath, ";", CurrentSheetName

End Sub

'**************************************************************
Private Sub copyDataFromCsvFileToSheet(parFileName As String, _
                                       parDelimiter As String, parSheetName As String)

    Dim Data As Variant                          'Array for the file values

    'Function call - the file is read into the array
    Data = getDataFromFile(parFileName, parDelimiter)

    'If the array isn't empty it is inserted into
    'the sheet in one swift operation.
    If Not isArrayEmpty(Data) Then
        'If you want to operate directly on the array,
        'you can leave out the following lines.
        With Sheets(parSheetName)
            'Delete any old content
            .Cells.ClearContents
            'A range gets the same dimensions as the array
            'and the array values are inserted in one operation.
            .Cells(1, 1).Resize(UBound(Data, 1), UBound(Data, 2)) = Data
        End With
    End If

End Sub

'**************************************************************
Public Function isArrayEmpty(parArray As Variant) As Boolean
    'Returns False if not an array or a dynamic array
    'that hasn't been initialised (ReDim) or
    'deleted (Erase).

    If IsArray(parArray) = False Then isArrayEmpty = True
    On Error Resume Next
    If UBound(parArray) < LBound(parArray) Then
        isArrayEmpty = True
        Exit Function
    Else
        isArrayEmpty = False
    End If

End Function

'**************************************************************
Private Function getDataFromFile(parFileName As String, _
                                 parDelimiter As String, _
                                 Optional parExcludeCharacter As String = "") As Variant
    'parFileName is the delimited file (csv, txt ...)
    'parDelimiter is the separator, e.g. semicolon.
    'The function returns an empty array, if the file
    'is empty or cannot be opened.
    'Number of columns is based on the line with most
    'columns and not the first line.
    'parExcludeCharacter: Some csv files have strings in
    'quotations marks ("ABC"), and if parExcludeCharacter = """"
    'quotation marks are removed.

    Dim locLinesList() As Variant                'Array
    Dim locData As Variant                       'Array
    Dim i As Long                                'Counter
    Dim j As Long                                'Counter
    Dim locNumRows As Long                       'Nb of rows
    Dim locNumCols As Long                       'Nb of columns
    Dim fso As Variant                           'File system object
    Dim ts As Variant                            'File variable
    Const REDIM_STEP = 10000                     'Constant

    'If this fails you need to reference Microsoft Scripting Runtime.
    'You select this in "Tools" (VBA editor menu).
    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error GoTo error_open_file
    'Sets ts = the file
    Set ts = fso.OpenTextFile(parFileName)
    On Error GoTo unhandled_error

    'Initialise the array
    ReDim locLinesList(1 To 1) As Variant
    i = 0
    'Loops through the file, counts the number of lines (rows)
    'and finds the highest number of columns.
    Do While Not ts.AtEndOfStream
        'If the row number Mod 10000 = 0
        'we redimension the array.
        If i Mod REDIM_STEP = 0 Then
            ReDim Preserve locLinesList _
                  (1 To UBound(locLinesList, 1) + REDIM_STEP) As Variant
        End If
        locLinesList(i + 1) = Split(ts.ReadLine, parDelimiter)
        j = UBound(locLinesList(i + 1), 1)       'Nb of columns in present row
        'If the number of columns is then highest so far.
        'the new number is saved.
        If locNumCols < j Then locNumCols = j
        i = i + 1
    Loop

    ts.Close                                     'Close file

    locNumRows = i

    'If number of rows is zero
    If locNumRows = 0 Then Exit Function

    ReDim locData(1 To locNumRows, 1 To locNumCols + 1) As Variant

    'Copies the file values into an array.
    'If parExcludeCharacter has a value,
    'the characters are removed.
    If parExcludeCharacter <> "" Then
        For i = 1 To locNumRows
            For j = 0 To UBound(locLinesList(i), 1)
                If Left(locLinesList(i)(j), 1) = parExcludeCharacter Then
                    If Right(locLinesList(i)(j), 1) = parExcludeCharacter Then
                        locLinesList(i)(j) = _
                                           Mid(locLinesList(i)(j), 2, Len(locLinesList(i)(j)) - 2)
                    Else
                        locLinesList(i)(j) = _
                                           Right(locLinesList(i)(j), Len(locLinesList(i)(j)) - 1)
                    End If
                ElseIf Right(locLinesList(i)(j), 1) = parExcludeCharacter Then
                    locLinesList(i)(j) = _
                                       Left(locLinesList(i)(j), Len(locLinesList(i)(j)) - 1)
                End If
                locData(i, j + 1) = locLinesList(i)(j)
            Next j
        Next i
    Else
        For i = 1 To locNumRows
            For j = 0 To UBound(locLinesList(i), 1)
                locData(i, j + 1) = locLinesList(i)(j)
            Next j
        Next i
    End If

    getDataFromFile = locData

    Exit Function

    error_open_file:                             'Returns empty Variant
    unhandled_error:                             'Returns empty Variant

End Function
