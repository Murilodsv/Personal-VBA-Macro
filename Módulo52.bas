Attribute VB_Name = "Módulo52"
Sub Macro32()
Attribute Macro32.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro32 Macro
'

'
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;F:\Output_Yg_3e\RNDPOSTERIOR_macT.OUT", Destination:=Range("$A$1"))
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
    Range("E4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Importa").Select
    Columns("A:T").Select
    Range("A346").Activate
    Application.CutCopyMode = False
    Selection.QueryTable.Delete
    Selection.ClearContents
End Sub
Sub Macro33()
Attribute Macro33.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro33 Macro
'

'
    ActiveWindow.ScrollColumn = 2
    Range("T4").Select
    ActiveWindow.SmallScroll ToRight:=39
    Range("BA1").Select
    ActiveCell.FormulaR1C1 = "=COUNTA(C[-52])"
    Range("BA1").Select
    Selection.ClearContents
End Sub
