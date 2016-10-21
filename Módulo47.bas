Attribute VB_Name = "Módulo47"
Sub Macro28()
Attribute Macro28.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro28 Macro
'

'
For x = 34 To 37

Windows("lista.xlsx").Activate
wth = Sheets("lista").Range("A" & x).Value
AMP = Sheets("lista").Range("C" & x).Value

    Workbooks.Add
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\Murilo\DOUTORADO\AGMIP\DataSubmission_Stage1_v2\WTH_DSSAT\" & wth & "" _
        , Destination:=Range("$A$1"))
        .Name = "" & wth & ""
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
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    Rows("4:4").Select
    Selection.Replace What:="-99.0 -99.0 -99.0", Replacement:=" " & AMP & " -99.0 -99.0", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("A1").Select
    
ActiveWorkbook.SaveAs Filename:="C:\Murilo\DOUTORADO\AGMIP\DataSubmission_Stage1_v2\WTH_DSSAT\AMP\" & wth & "", _
       FileFormat:=xlTextPrinter, CreateBackup:=False

Application.DisplayAlerts = False
ActiveWindow.Close

Next

End Sub
