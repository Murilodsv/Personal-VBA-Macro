Attribute VB_Name = "Módulo51"
Sub RNDPOSTERIOR_OUT()
Attribute RNDPOSTERIOR_OUT.VB_ProcData.VB_Invoke_Func = " \n14"
'
'
'

dire = ThisWorkbook.Path

For x = 1 To 108

arq = Sheets("Lista").Range("A" & x + 3).Value
cen = Sheets("Lista").Range("B" & x + 3).Value

    Sheets("Importa").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & dire & "\" & cen & "\" & arq & ".OUT", Destination:=Range("$A$1"))
        .Name = "RNDPOSTERIOR_macT"
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
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    Range("A1").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("Lista").Select
    Range(Cells(x + 3, 5), Cells(x + 3, 5)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets("Importa").Select
    Columns("A:T").Select
    Application.DisplayAlerts = False
    Selection.ClearContents

Next

Sheets("Lista").Select

End Sub
