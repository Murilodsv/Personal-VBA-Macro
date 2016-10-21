Attribute VB_Name = "Módulo18"
Sub WTH_TO_EXCEL()
Attribute WTH_TO_EXCEL.VB_ProcData.VB_Invoke_Func = " \n14"
'
' WTH_TO_EXCEL Macro
'

'
Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\estacoes_selecao.xlsx"

For x = 1 To 30

Windows("estacoes_selecao.xlsx").Activate
wth = Sheets("estacoes_selecao").Range("AU" & x + 1).Value
ANO1 = Sheets("estacoes_selecao").Range("AV" & x + 1).Value

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\Interpolado\WTH\IMPORTA.xlsx"

    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\DSSAT45\Weather\" & wth & "" & ANO1 & "01.WTH", Destination:=Range(Cells(1, 1), Cells(1, 1)))
        .Name = wth
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
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    linha = 371
    
    For y = 1 To 32
    
    ano = ANO1 + y
    
    If ano >= 100 And ano < 110 Then
    
    ano = "0" & ANO1 + y - 100
    
    ElseIf ano >= 110 Then
    
    ano = ANO1 + y - 100
    
    Else
    
    ano = ANO1 + y
    
    End If
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\DSSAT45\Weather\" & wth & "" & ano & "01.WTH", Destination:=Range(Cells(linha, 1), Cells(linha, 1)))
        .Name = wth
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
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    Range(Cells(linha, 1), Cells(linha + 4, 10)).Select
    Selection.Delete Shift:=xlUp
    
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=(RC[-1]/4)-(TRUNC(RC[-1]/4))"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = ANO1 + y
    
    Calculate
    
    ANOB = Range("N2").Value
    Range("N2:M2").Select
    Selection.ClearContents
    
    If ANOB = 0 Then
    
    linha = linha + 366
    
    Else
    
    linha = linha + 365
    
    End If
    
    Next
    
    
    Windows("IMPORTA.xlsx").Activate
    ActiveWorkbook.SaveAs Filename:= _
        "C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\Interpolado\WTH\" & wth & ".xls" _
        , FileFormat:=xlExcel8, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
Application.DisplayAlerts = False
ActiveWindow.Close
    
    
    
    Next
    
    Call BH_WTH
    
End Sub
Sub BH_WTH()
Attribute BH_WTH.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro17 Macro
'

'
Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\BH_ESTACOES.xlsx"
Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\estacoes_selecao.xlsx"

linha = 1
LINHAD = 1

For x = 1 To 30

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\MODELO_ANALISE_SERIE_WTH.xlsx"

Application.Calculation = xlManual

Windows("estacoes_selecao.xlsx").Activate
wth = Sheets("estacoes_selecao").Range("AU" & x + 1).Value
Lat = Sheets("estacoes_selecao").Range("A" & x + 1).Value
Lon = Sheets("estacoes_selecao").Range("B" & x + 1).Value
nome = Sheets("estacoes_selecao").Range("AY" & x + 1).Value

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\Interpolado\WTH\" & wth & ".xls"

Range("C6").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("MODELO_ANALISE_SERIE_WTH.xlsx").Activate
Sheets("DIA").Select
Range("B7").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & wth & ".xls").Activate
Range("B6").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("MODELO_ANALISE_SERIE_WTH.xlsx").Activate
Sheets("DIA").Select
Range("E7").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("" & wth & ".xls").Activate
Range("B4:E4").Select
Selection.Copy

Windows("MODELO_ANALISE_SERIE_WTH.xlsx").Activate
Range("B1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
        
Calculate
       
Sheets("BH Mensal").Select
ActiveSheet.ChartObjects("Gráfico 4").Activate
ActiveChart.ChartArea.Copy

Windows("BH_ESTACOES.xlsx").Activate
Sheets("BH_MENSAL").Select
Range(Cells(linha + 1, 1), Cells(linha + 1, 1)).Select
ActiveSheet.Pictures.Paste.Select

Range(Cells(linha, 1), Cells(linha, 1)).Select
ActiveCell.FormulaR1C1 = "" & wth & " - " & nome & ""

Windows("MODELO_ANALISE_SERIE_WTH.xlsx").Activate
Sheets("BH Mensal").Select
ActiveSheet.ChartObjects("Gráfico 7").Activate
ActiveChart.ChartArea.Copy

Windows("BH_ESTACOES.xlsx").Activate
Sheets("BH_MENSAL").Select
Range(Cells(linha + 1, 11), Cells(linha + 1, 11)).Select
ActiveSheet.Pictures.Paste.Select

Windows("MODELO_ANALISE_SERIE_WTH.xlsx").Activate
Sheets("BH Mensal").Select
ActiveSheet.ChartObjects("Gráfico 3").Activate
ActiveChart.ChartArea.Copy

Windows("BH_ESTACOES.xlsx").Activate
Sheets("BH_MENSAL").Select
Range(Cells(linha + 1, 21), Cells(linha + 1, 21)).Select
ActiveSheet.Pictures.Paste.Select

Windows("MODELO_ANALISE_SERIE_WTH.xlsx").Activate
Sheets("BH Mensal").Select
ActiveSheet.ChartObjects("Gráfico 11").Activate
ActiveChart.ChartArea.Copy

Windows("BH_ESTACOES.xlsx").Activate
Sheets("BH_MENSAL").Select
Range(Cells(linha + 1, 31), Cells(linha + 1, 31)).Select
ActiveSheet.Pictures.Paste.Select

'-----------------------------------------------------------------------


Windows("MODELO_ANALISE_SERIE_WTH.xlsx").Activate
Sheets("BH Decendial").Select
ActiveSheet.ChartObjects("Gráfico 23").Activate
ActiveChart.ChartArea.Copy

Windows("BH_ESTACOES.xlsx").Activate
Sheets("BH_DECENDIAL").Select
Range(Cells(LINHAD + 1, 1), Cells(LINHAD + 1, 1)).Select
ActiveSheet.Pictures.Paste.Select

Range(Cells(LINHAD, 1), Cells(LINHAD, 1)).Select
ActiveCell.FormulaR1C1 = "" & wth & " - " & nome & ""

Windows("MODELO_ANALISE_SERIE_WTH.xlsx").Activate
Sheets("BH Decendial").Select
ActiveSheet.ChartObjects("Gráfico 22").Activate
ActiveChart.ChartArea.Copy

Windows("BH_ESTACOES.xlsx").Activate
Sheets("BH_DECENDIAL").Select
Range(Cells(LINHAD + 1, 12), Cells(LINHAD + 1, 12)).Select
ActiveSheet.Pictures.Paste.Select

Windows("MODELO_ANALISE_SERIE_WTH.xlsx").Activate
Sheets("BH Decendial").Select
ActiveSheet.ChartObjects("Gráfico 24").Activate
ActiveChart.ChartArea.Copy

Windows("BH_ESTACOES.xlsx").Activate
Sheets("BH_DECENDIAL").Select
Range(Cells(LINHAD + 1, 23), Cells(LINHAD + 1, 23)).Select
ActiveSheet.Pictures.Paste.Select

Windows("MODELO_ANALISE_SERIE_WTH.xlsx").Activate
Sheets("BH Decendial").Select
ActiveSheet.ChartObjects("Gráfico 26").Activate
ActiveChart.ChartArea.Copy

Windows("BH_ESTACOES.xlsx").Activate
Sheets("BH_DECENDIAL").Select
Range(Cells(LINHAD + 1, 34), Cells(LINHAD + 1, 34)).Select
ActiveSheet.Pictures.Paste.Select

Windows("MODELO_ANALISE_SERIE_WTH.xlsx").Activate
Sheets("SINTESE").Select
ActiveSheet.ChartObjects("Gráfico 2").Activate
ActiveChart.ChartArea.Copy

Windows("BH_ESTACOES.xlsx").Activate
Sheets("CLIMA").Select
Range(Cells(linha + 1, 1), Cells(linha + 1, 1)).Select
ActiveSheet.Pictures.Paste.Select

Range(Cells(linha, 1), Cells(linha, 1)).Select
ActiveCell.FormulaR1C1 = "" & wth & " - " & nome & ""

Windows("MODELO_ANALISE_SERIE_WTH.xlsx").Activate
Sheets("SINTESE").Select
ActiveSheet.ChartObjects("Gráfico 3").Activate
ActiveChart.ChartArea.Copy

Windows("BH_ESTACOES.xlsx").Activate
Sheets("CLIMA").Select
Range(Cells(linha + 1, 12), Cells(linha + 1, 12)).Select
ActiveSheet.Pictures.Paste.Select

Windows("MODELO_ANALISE_SERIE_WTH.xlsx").Activate
Sheets("BH Sequencial").Select
Range("AM19:BU42").Select
Selection.Copy

Windows("BH_ESTACOES.xlsx").Activate
Sheets("BH_SEQUENCIAL").Select
Range(Cells(LINHAD + 1, 1), Cells(LINHAD + 1, 1)).Select
ActiveSheet.Pictures.Paste.Select

'MEDIAS-----------------------

Windows("MODELO_ANALISE_SERIE_WTH.xlsx").Activate
Sheets("DECENDIO").Select
Range("B39:ES39").Select
Selection.Copy

Windows("BH_ESTACOES.xlsx").Activate
Sheets("MEDIA_DECENDIAL").Select
Range(Cells(x + 2, 2), Cells(x + 2, 2)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("MODELO_ANALISE_SERIE_WTH.xlsx").Activate
Sheets("DECENDIO").Select
Range("B40:ES40").Select
Selection.Copy

Windows("BH_ESTACOES.xlsx").Activate
Sheets("MEDIA_DECENDIAL").Select
Range(Cells(x + 37, 2), Cells(x + 37, 2)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("MODELO_ANALISE_SERIE_WTH.xlsx").Activate
Sheets("MES").Select
Range("B39:ES39").Select
Selection.Copy

Windows("BH_ESTACOES.xlsx").Activate
Sheets("MEDIA_MENSAL").Select
Range(Cells(x + 2, 2), Cells(x + 2, 2)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("MODELO_ANALISE_SERIE_WTH.xlsx").Activate
Sheets("MES").Select
Range("B40:ES40").Select
Selection.Copy

Windows("BH_ESTACOES.xlsx").Activate
Sheets("MEDIA_MENSAL").Select
Range(Cells(x + 37, 2), Cells(x + 37, 2)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False


Windows("MODELO_ANALISE_SERIE_WTH.xlsx").Activate
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\Interpolado\WTH\" & wth & "_SINTESE.xlsx", _
FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
ActiveWindow.Close

Windows("" & wth & ".xls").Activate
Application.DisplayAlerts = False
ActiveWindow.Close

linha = linha + 23
LINHAD = LINHAD + 29

Next

End Sub
