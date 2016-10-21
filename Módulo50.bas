Attribute VB_Name = "Módulo50"
Sub Macro31()
Attribute Macro31.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro31 Macro
'

'


    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\DSSAT45\Seasonal\ESSP1401.OPG", Destination:=Range("$E$1"))
        .Name = "ESSP1401"
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
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Calculate
    Columns("E:BG").Select
    Selection.QueryTable.Delete
    Selection.ClearContents
    Selection.Delete Shift:=xlToLeft
    Range("E1").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\DSSAT45\Seasonal\ESSP1401.OPG", Destination:=Range("$E$1"))
        .Name = "ESSP1401_1"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlOverwriteCells
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
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Calculate
    ActiveWindow.ScrollRow = 87
    ActiveWindow.ScrollRow = 519
    ActiveWindow.ScrollRow = 605
    ActiveWindow.ScrollRow = 691
    ActiveWindow.ScrollRow = 777
    ActiveWindow.ScrollRow = 864
    ActiveWindow.ScrollRow = 950
    ActiveWindow.ScrollRow = 1036
    ActiveWindow.ScrollRow = 950
    ActiveWindow.ScrollRow = 864
    ActiveWindow.ScrollRow = 691
    ActiveWindow.ScrollRow = 605
    ActiveWindow.ScrollRow = 519
    ActiveWindow.ScrollRow = 432
    ActiveWindow.ScrollRow = 346
    ActiveWindow.ScrollRow = 519
    ActiveWindow.ScrollRow = 605
    ActiveWindow.ScrollRow = 691
    ActiveWindow.ScrollRow = 777
    ActiveWindow.SmallScroll Down:=-999
    Range("B14").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.AutoFilter
    ActiveSheet.Range("$B$14:$B$40392").AutoFilter Field:=1, Criteria1:="<>", _
        Operator:=xlAnd
    ActiveSheet.Range("$B$14:$B$40392").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SmallScroll Down:=-96
End Sub
