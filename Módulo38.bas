Attribute VB_Name = "Módulo38"
Sub organiza_estacoes_ana()
'
'
'

'

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\estacoes_selecao.xlsx"

For x = 1 To 30

Windows("estacoes_selecao.xlsx").Activate

BH = Sheets("estacoes_selecao").Range("AU" & x + 1).Value

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\Interpolado\WTH\" & BH & "_SINTESE.xlsx"

Sheets("BH Mensal").Select
Range("W21").Select
ActiveCell.FormulaR1C1 = "=SUM(RC[2]:RC[3])"
Range("W21").Select
Selection.AutoFill Destination:=Range("W21:W32")
Range("W21:W32").Select
Calculate
Selection.Copy

Windows("estacoes_selecao.xlsx").Activate
Sheets("BH").Select

Range(Cells(x + 1, 2), Cells(x + 1, 2)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True

Windows("" & BH & "_SINTESE.xlsx").Activate
Application.DisplayAlerts = False
ActiveWindow.Close

Next

End Sub


Sub bh_decendial()
'
'
'

'

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\BH_ESTACOES_7_2.xlsx"

For x = 1 To 30

Windows("BH_ESTACOES_7_2.xlsx").Activate

BH = Sheets("MEDIA_MENSAL").Range("A" & x + 2).Value

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\Interpolado\WTH\" & BH & "_SINTESE.xlsx"

Windows("BH_ESTACOES_7_2.xlsx").Activate
Sheets("BH_medio_dec").Select
Range("A2:A37").Select
Selection.Copy

Windows("" & BH & "_SINTESE.xlsx").Activate
Sheets("BH Sequencial").Select
Range("AN50").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Range("AO50").Select
    ActiveCell.FormulaR1C1 = "=AVERAGEIF(C25,RC40,C[-5])"
    Range("AO50").Select
    Selection.AutoFill Destination:=Range("AO50:AP50"), Type:=xlFillDefault
    Range("AO50:AP50").Select
    Selection.AutoFill Destination:=Range("AO50:AP85")
    Range("AO50:AP85").Select
    Calculate

Range("AO50:AO85").Select
Selection.Copy

Windows("BH_ESTACOES_7_2.xlsx").Activate
Range(Cells(2, x + 1), Cells(2, x + 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("" & BH & "_SINTESE.xlsx").Activate
Range("AP50:AP85").Select
Selection.Copy

Windows("BH_ESTACOES_7_2.xlsx").Activate
Range(Cells(2, x + 40), Cells(2, x + 40)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("" & BH & "_SINTESE.xlsx").Activate
Application.DisplayAlerts = False
ActiveWindow.Close

Next

End Sub


