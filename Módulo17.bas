Attribute VB_Name = "Módulo17"
Sub Consistencia()
Attribute Consistencia.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro16 Macro
'

'

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\estacoes_selecao.xlsx"

For x = 1 To 30

Windows("estacoes_selecao.xlsx").Activate
wth = Sheets("estacoes_selecao").Range("AU" & x + 1).Value

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\Interpolado\WTH\" & wth & ".xls"

    Range("J7").Select
    ActiveCell.FormulaR1C1 = "=LARGE(C[-8],1)"
    Range("J7").Select
    Selection.AutoFill Destination:=Range("J7:N7"), Type:=xlFillDefault
    Range("O7").Select
    ActiveCell.FormulaR1C1 = "=SMALL(C[-13],1)"
    Range("O7").Select
    Selection.AutoFill Destination:=Range("O7:S7"), Type:=xlFillDefault
    Range("O7:S7").Select
    Rows("1:5").Select
    Selection.ClearContents
    Columns("A:F").Select
    ' Selection.Replace What:="-99", Replacement:="", LookAt:=xlWhole, _
      '  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
       ' ReplaceFormat:=False
        

    Range("J7:S7").Select
    Selection.Copy

    Windows("estacoes_selecao.xlsx").Activate
    Sheets("Consistencia").Select
    Range(Cells(x + 2, 12), Cells(x + 2, 12)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Windows("" & wth & ".xls").Activate
    Application.DisplayAlerts = False
    ActiveWindow.Close


Next

End Sub
