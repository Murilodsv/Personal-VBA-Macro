Attribute VB_Name = "Módulo64"
Sub Macro47()
Attribute Macro47.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro47 Macro
'

'
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\Murilo\DOUTORADO\Datalogger\Dados\CR23X_final_storage_1_01_09_2014.dat" _
        , Destination:=Range("$B$13"))
        .Name = "CR23X_final_storage_1_01_09_2014"
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
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.QueryTable.Delete
    Selection.ClearContents
    Range("A12").Select
    Columns("B:B").EntireColumn.AutoFit
    Range("B5").Select
End Sub
