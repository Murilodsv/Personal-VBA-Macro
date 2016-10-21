Attribute VB_Name = "Módulo1"
Sub copiaecola()


Workbooks.Open Filename:="C:\Murilo\DSSAT\ARTIGO\SINTESE.xlsx"

linha = 5
LINHAFIM = 53

For x = 1 To 39

Workbooks.Open Filename:="C:\Murilo\DSSAT\ARTIGO\MODELO.xlsx"

Windows("SINTESE.xlsx").Activate
arquivo = Sheets("Plan1").Range("K" & x + 1).Value

For y = 1 To 6

Windows("SINTESE.xlsx").Activate
nome = Sheets("Plan1").Range("J" & y + 1).Value
plan = Sheets("Plan1").Range("H" & y + 1).Value

Workbooks.Open Filename:="C:\Murilo\DSSAT\ARTIGO\" & nome & ""

Sheets("RESULTADO").Select
Range(Cells(linha, 1), Cells(LINHAFIM, 17)).Select
Selection.Copy

Windows("MODELO.xlsx").Activate
Sheets("" & plan & "").Select
Range("A5").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
Windows("" & nome & "").Activate
Range(Cells(linha, 19), Cells(LINHAFIM, 53)).Select
Selection.Copy

Windows("MODELO.xlsx").Activate
Sheets("" & plan & "").Select
Range("S5").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & nome & "").Activate
Application.DisplayAlerts = False
ActiveWindow.Close


Next

Windows("MODELO.xlsx").Activate
NOMEMODELO = Sheets("LAT_BASELINE").Range("D" & 5).Value

ActiveWorkbook.SaveAs Filename:= _
        "C:\Murilo\DSSAT\ARTIGO\ANALISE\" & NOMEMODELO & "", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
        ActiveWindow.Close

linha = linha + 49
LINHAFIM = LINHAFIM + 49

Next

End Sub

Sub Macro3()
'
' Macro3 Macro
'

'

Workbooks.Open Filename:="C:\Murilo\DSSAT\ARTIGO\SINTESE.xlsx"

K = 1
For x = 1 To 3

Workbooks.Open Filename:="C:\Murilo\DSSAT\ARTIGO\ACUMULADO.xlsx"


For y = 1 To 13

Windows("SINTESE.xlsx").Activate
arquivo = Sheets("Plan1").Range("O" & K + 1).Value
CORTE = Sheets("Plan1").Range("M" & K + 1).Value
Workbooks.Open Filename:="C:\Murilo\DSSAT\ARTIGO\ANALISE\" & arquivo & ""

Windows("" & arquivo & "").Activate
Sheets("LAT_BASELINE").Select

Range("D60:D77").Select
Selection.Copy

Windows("ACUMULADO.xlsx").Activate
Sheets("LAT_BASELINE").Select
Range(Cells(2, y + 1), Cells(2, y + 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & arquivo & "").Activate
Sheets("LAT_A2").Select

Range("D60:D77").Select
Selection.Copy

Windows("ACUMULADO.xlsx").Activate
Sheets("LAT_A2").Select
Range(Cells(2, y + 1), Cells(2, y + 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & arquivo & "").Activate
Sheets("LAT_B2").Select

Range("D60:D77").Select
Selection.Copy

Windows("ACUMULADO.xlsx").Activate
Sheets("LAT_B2").Select
Range(Cells(2, y + 1), Cells(2, y + 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("" & arquivo & "").Activate
Sheets("POD_BASELINE").Select

Range("D60:D77").Select
Selection.Copy

Windows("ACUMULADO.xlsx").Activate
Sheets("POD_BASELINE").Select
Range(Cells(2, y + 1), Cells(2, y + 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & arquivo & "").Activate
Sheets("POD_A2").Select

Range("D60:D77").Select
Selection.Copy

Windows("ACUMULADO.xlsx").Activate
Sheets("POD_A2").Select
Range(Cells(2, y + 1), Cells(2, y + 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & arquivo & "").Activate
Sheets("POD_B2").Select

Range("D60:D77").Select
Selection.Copy

Windows("ACUMULADO.xlsx").Activate
Sheets("POD_B2").Select
Range(Cells(2, y + 1), Cells(2, y + 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("" & arquivo & "").Activate
Sheets("MEDIA_BASELINE").Select

Range("D60:D77").Select
Selection.Copy

Windows("ACUMULADO.xlsx").Activate
Sheets("MEDIA_BASELINE").Select
Range(Cells(2, y + 1), Cells(2, y + 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & arquivo & "").Activate
Sheets("MEDIA_A2").Select

Range("D60:D77").Select
Selection.Copy

Windows("ACUMULADO.xlsx").Activate
Sheets("MEDIA_A2").Select
Range(Cells(2, y + 1), Cells(2, y + 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & arquivo & "").Activate
Sheets("MEDIA_B2").Select

Range("D60:D77").Select
Selection.Copy

Windows("ACUMULADO.xlsx").Activate
Sheets("MEDIA_B2").Select
Range(Cells(2, y + 1), Cells(2, y + 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        

'------------------------------------------------------------------------------------

Windows("SINTESE.xlsx").Activate
arquivo = Sheets("Plan1").Range("O" & K + 1).Value
Workbooks.Open Filename:="C:\Murilo\DSSAT\ARTIGO\ANALISE\" & arquivo & ""

Windows("" & arquivo & "").Activate
Sheets("LAT_BASELINE").Select

Range("F60:F67").Select
Selection.Copy

Windows("ACUMULADO.xlsx").Activate
Sheets("LAT_BASELINE").Select
Range(Cells(22, y + 1), Cells(22, y + 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & arquivo & "").Activate
Sheets("LAT_A2").Select

Range("F60:F67").Select
Selection.Copy

Windows("ACUMULADO.xlsx").Activate
Sheets("LAT_A2").Select
Range(Cells(22, y + 1), Cells(22, y + 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & arquivo & "").Activate
Sheets("LAT_B2").Select

Range("F60:F67").Select
Selection.Copy

Windows("ACUMULADO.xlsx").Activate
Sheets("LAT_B2").Select
Range(Cells(22, y + 1), Cells(22, y + 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("" & arquivo & "").Activate
Sheets("POD_BASELINE").Select

Range("F60:F67").Select
Selection.Copy

Windows("ACUMULADO.xlsx").Activate
Sheets("POD_BASELINE").Select
Range(Cells(22, y + 1), Cells(22, y + 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & arquivo & "").Activate
Sheets("POD_A2").Select

Range("F60:F67").Select
Selection.Copy

Windows("ACUMULADO.xlsx").Activate
Sheets("POD_A2").Select
Range(Cells(22, y + 1), Cells(22, y + 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & arquivo & "").Activate
Sheets("POD_B2").Select

Range("F60:F67").Select
Selection.Copy

Windows("ACUMULADO.xlsx").Activate
Sheets("POD_B2").Select
Range(Cells(22, y + 1), Cells(22, y + 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("" & arquivo & "").Activate
Sheets("MEDIA_BASELINE").Select

Range("F60:F67").Select
Selection.Copy

Windows("ACUMULADO.xlsx").Activate
Sheets("MEDIA_BASELINE").Select
Range(Cells(22, y + 1), Cells(22, y + 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & arquivo & "").Activate
Sheets("MEDIA_A2").Select

Range("F60:F67").Select
Selection.Copy

Windows("ACUMULADO.xlsx").Activate
Sheets("MEDIA_A2").Select
Range(Cells(22, y + 1), Cells(22, y + 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & arquivo & "").Activate
Sheets("MEDIA_B2").Select

Range("F60:F67").Select
Selection.Copy

Windows("ACUMULADO.xlsx").Activate
Sheets("MEDIA_B2").Select
Range(Cells(22, y + 1), Cells(22, y + 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False


Windows("" & arquivo & "").Activate
Application.DisplayAlerts = False
ActiveWindow.Close
K = K + 1

Next

Windows("ACUMULADO.xlsx").Activate
Sheets("MEDIA_BASELINE").Select
Range("O1:AA29").Select
Selection.Copy

Range("B1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
Sheets("MEDIA_A2").Select
Range("O1:AA29").Select
Selection.Copy

Range("B1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Sheets("MEDIA_B2").Select
Range("O1:AA29").Select
Selection.Copy

Range("B1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Sheets("LAT_BASELINE").Select
Range("O1:AA29").Select
Selection.Copy

Range("B1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Sheets("LAT_A2").Select
Range("O1:AA29").Select
Selection.Copy

Range("B1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Sheets("LAT_B2").Select
Range("O1:AA29").Select
Selection.Copy

Range("B1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Sheets("POD_BASELINE").Select
Range("O1:AA29").Select
Selection.Copy

Range("B1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Sheets("POD_A2").Select
Range("O1:AA29").Select
Selection.Copy

Range("B1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Sheets("POD_B2").Select
Range("O1:AA29").Select
Selection.Copy

Range("B1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False


Windows("ACUMULADO.xlsx").Activate
ActiveWorkbook.SaveAs Filename:= _
        "C:\Murilo\DSSAT\ARTIGO\ANALISE\ACUMULADO_" & CORTE & "", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
        ActiveWindow.Close

    
Next

    
End Sub


Sub AREA_ARC()
Attribute AREA_ARC.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    
Workbooks.Open Filename:="C:\Murilo\ARCGIS\FGV\PERSON\GIRASSOL\GIRASSOL_MINISTERIO\AREAS_ZAE_CANA_GIRASSOL.xlsx"

    
    linha = 2
    For x = 8 To 21
    
    Windows("AREAS_ZAE_CANA_GIRASSOL.xlsx").Activate
    
    uf = Sheets("PLAN2").Range("P" & x).Value
    
    Workbooks.Open Filename:="C:\Murilo\ARCGIS\FGV\PERSON\ZAECANA_GIRASSOL\MAPA\TABELAS\GIRASSOL_ZAECANA_" & uf & ".csv"
    
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-15],""<>"")"
    
    NLINHA = Range("P1").Value - 1
    
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("AREAS_ZAE_CANA_GIRASSOL.xlsx").Activate
    Sheets("ZAE_CANA_GIRASSOL").Select
    Range(Cells(linha, 1), Cells(linha, 1)).Select
    ActiveSheet.Paste
    
    Windows("GIRASSOL_ZAECANA_" & uf & ".csv").Activate

    Application.DisplayAlerts = False
    ActiveWindow.Close

    
    linha = linha + NLINHA
    
    Next
End Sub
Sub AGRUPA_SERIES_PIRA()
Attribute AGRUPA_SERIES_PIRA.VB_ProcData.VB_Invoke_Func = " \n14"
'
' AGRUPA SERIES
'

'

Workbooks.Open Filename:="C:\Murilo\MESTRADO\Dados_met\Dados Diarios Pira\AUTOMATICA\DADOS_DIARIOS_AUTOMATICA_PIRA.xlsx"


For x = 1 To 15

Windows("DADOS_DIARIOS_AUTOMATICA_PIRA.xlsx").Activate
arq = Sheets("Lista").Range("A" & x).Value

Workbooks.Open Filename:="C:\Murilo\MESTRADO\Dados_met\Dados Diarios Pira\AUTOMATICA\" & arq & ".xls"
    Cells.Replace What:="---------", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    Cells.Replace What:=" ", Replacement:="", LookAt:=xlWhole, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("DADOS_DIARIOS_AUTOMATICA_PIRA.xlsx").Activate
    
    linha = Range("S1").Value + 9
    
    Range(Cells(linha, 1), Cells(linha, 1)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

For y = 1 To 11

Windows("" & arq & ".xls").Activate
 
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select

    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("DADOS_DIARIOS_AUTOMATICA_PIRA.xlsx").Activate
    
    linha = Range("S1").Value + 9
    
    Range(Cells(linha, 1), Cells(linha, 1)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   
    
    Next
    
    Windows("" & arq & ".xls").Activate
    Application.DisplayAlerts = False
    ActiveWindow.Close

    
    
    
    Next
End Sub
