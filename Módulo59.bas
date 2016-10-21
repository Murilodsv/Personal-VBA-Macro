Attribute VB_Name = "Módulo59"
Sub Macro37()
Attribute Macro37.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro37 Macro
'

'
    Sheets("IMPORTA").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\Murilo\MACRO\Fabio\Barra.txt", Destination:=Range("$A$1"))
        .Name = "Barra"
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
    ActiveWindow.SmallScroll Down:=-75
End Sub
