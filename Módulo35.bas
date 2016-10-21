Attribute VB_Name = "Módulo35"
Sub SRAD_BC_criaWTH()
Attribute SRAD_BC_criaWTH.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro23 Macro
'

'

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\Interpolado\BRISTOW_CAMPBEL_SRAD.xlsx"
Application.Calculation = xlManual

For x = 1 To 30

Windows("BRISTOW_CAMPBEL_SRAD.xlsx").Activate
wth = Sheets("LISTA").Range("A" & x).Value
EST = Sheets("LISTA").Range("B" & x).Value

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\Interpolado\WTH\" & wth & ".xls"
Range("B6:B12058").Select
Selection.Copy

Windows("BRISTOW_CAMPBEL_SRAD.xlsx").Activate
Sheets("BC").Select
Range("F7").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("" & wth & ".xls").Activate
Range("C6:D12058").Select
Selection.Copy

Windows("BRISTOW_CAMPBEL_SRAD.xlsx").Activate
Sheets("BC").Select
Range("C7").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & wth & ".xls").Activate
Range("E6:E12058").Select
Selection.Copy

Windows("BRISTOW_CAMPBEL_SRAD.xlsx").Activate
Sheets("BC").Select
Range("B7").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & wth & ".xls").Activate
Application.DisplayAlerts = False
ActiveWindow.Close

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\Interpolado\" & EST & ".xls"
Range("B1:B4").Select
Selection.Copy

Windows("BRISTOW_CAMPBEL_SRAD.xlsx").Activate
Sheets("BC").Select
Range("B1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & EST & ".xls").Activate
Range("E7:E12059").Select
Selection.Copy

Windows("BRISTOW_CAMPBEL_SRAD.xlsx").Activate
Sheets("BC").Select
Range("E7").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("" & EST & ".xls").Activate
Application.DisplayAlerts = False
ActiveWindow.Close

Windows("BRISTOW_CAMPBEL_SRAD.xlsx").Activate
Sheets("IMPORTA").Select

Columns("A:A").Select
Selection.ClearContents

    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\DSSAT45\Weather\" & wth & "0001.WTH", Destination:=Range("$A$1"))
        .Name = "ACCE0001"
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
    
Calculate

'-----cria WTH

For y = 1 To 33

Windows("BRISTOW_CAMPBEL_SRAD.xlsx").Activate
ano = Sheets("LISTA").Range("C" & y).Value

Sheets("IMPORTA").Select
Range("A6:A372").Select
Selection.ClearContents

Sheets("WTH_FINAL").Select
ActiveSheet.Range("$A$5:$A$12058").AutoFilter Field:=1, Criteria1:=ano

Range("O4").Select
Selection.End(xlDown).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Sheets("IMPORTA").Select
Range("A6").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Workbooks.Open Filename:="C:\Murilo\MACRO\IMPORTA.xlsx"

Windows("BRISTOW_CAMPBEL_SRAD.xlsx").Activate
Sheets("IMPORTA").Select
Columns("A:A").Select
Selection.Copy

Windows("IMPORTA.xlsx").Activate
Columns("A:A").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    If ano < 10 Then
    
    ano = "0" & ano & ""
    
    End If
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\Interpolado\WTH\NOVO_WTH\" & wth & "" & ano & "01.WTH", _
        FileFormat:=xlTextPrinter, CreateBackup:=False
    
    Application.DisplayAlerts = False
    ActiveWindow.Close
    
Next

Sheets("WTH_FINAL").Select
ActiveSheet.Range("$A$5:$A$12058").AutoFilter Field:=1

Next

End Sub
Sub criaWTH()
'
' Macro23 Macro
'

'

Workbooks.Open Filename:="C:\Murilo\MACRO\Cria_WTH.xlsx"
Application.Calculation = xlManual

For x = 1 To 6

Windows("Cria_WTH.xlsx").Activate
wth = Sheets("LISTA").Range("A" & x).Value
EST = Sheets("LISTA").Range("B" & x).Value

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\Interpolado\WTH\" & wth & ".xls"
Range("B6:B12058").Select
Selection.Copy

Windows("Cria_WTH.xlsx").Activate
Sheets("BC").Select
Range("F7").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("" & wth & ".xls").Activate
Range("C6:D12058").Select
Selection.Copy

Windows("Cria_WTH.xlsx").Activate
Sheets("BC").Select
Range("C7").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & wth & ".xls").Activate
Range("E6:E12058").Select
Selection.Copy

Windows("Cria_WTH.xlsx").Activate
Sheets("BC").Select
Range("B7").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & wth & ".xls").Activate
Application.DisplayAlerts = False
ActiveWindow.Close

Windows("Cria_WTH.xlsx").Activate
Sheets("IMPORTA").Select

Columns("A:A").Select
Selection.ClearContents

    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\DSSAT45\Weather\" & EST & "0001.WTH", Destination:=Range("$A$1"))
        .Name = "ACCE0001"
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
    
Calculate

'-----cria WTH

For y = 1 To 33

Windows("Cria_WTH.xlsx").Activate
ano = Sheets("LISTA").Range("C" & y).Value

Sheets("IMPORTA").Select
Range("A6:A372").Select
Selection.ClearContents

Sheets("WTH_FINAL").Select
ActiveSheet.Range("$A$5:$A$12058").AutoFilter Field:=1, Criteria1:=ano

Range("O4").Select
Selection.End(xlDown).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Sheets("IMPORTA").Select
Range("A6").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Workbooks.Open Filename:="C:\Murilo\MACRO\IMPORTA.xlsx"

Windows("Cria_WTH.xlsx").Activate
Sheets("IMPORTA").Select
Columns("A:A").Select
Selection.Copy

Windows("IMPORTA.xlsx").Activate
Columns("A:A").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    If ano < 10 Then
    
    ano = "0" & ano & ""
    
    End If
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\Interpolado\WTH\NOVO_WTH\NE\" & EST & "" & ano & "01.WTH", _
        FileFormat:=xlTextPrinter, CreateBackup:=False
    
    Application.DisplayAlerts = False
    ActiveWindow.Close
    
Next

Sheets("WTH_FINAL").Select
ActiveSheet.Range("$A$5:$A$12058").AutoFilter Field:=1

Next

End Sub

