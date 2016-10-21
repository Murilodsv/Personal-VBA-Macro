Attribute VB_Name = "Módulo16"
Sub CONTA_N()
Attribute CONTA_N.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro18 Macro
'

'
For x = 1 To 30
Windows("conta_n.xlsx").Activate
wth = Sheets("plan1").Range("A" & x + 1).Value

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\Interpolado\" & wth & "_merge_Rad_int.xls"

Windows("conta_n.xlsx").Activate
Range("H7:J7").Select
Selection.Copy

Windows("" & wth & "_merge_Rad_int.xls").Activate
Range("H7:J7").Select
ActiveSheet.Paste

Calculate

Range("H7:J7").Select
Selection.Copy

Windows("conta_n.xlsx").Activate
Range(Cells(x + 1, 2), Cells(x + 1, 2)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
 
Windows("" & wth & "_merge_Rad_int.xls").Activate
Application.DisplayAlerts = False
ActiveWindow.Close

Next
   
End Sub
Sub Conta_99()
Attribute Conta_99.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro19 Macro
'

'

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\estacoes_selecao.xlsx"

For x = 1 To 30
    
Windows("estacoes_selecao.xlsx").Activate
INMET = Sheets("estacoes_selecao").Range("D" & x + 1).Value

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\Interpolado\" & INMET & "_merge_Rad_int.xls"
   
    Range("H6").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-6],-99)"
    Range("H6").Select
    Selection.AutoFill Destination:=Range("H6:L6"), Type:=xlFillDefault
    Range("H6:L6").Select
    Selection.Copy
    

Windows("estacoes_selecao.xlsx").Activate
Range(Cells(x + 1, 37), Cells(x + 1, 37)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & INMET & "_merge_Rad_int.xls").Activate
Application.DisplayAlerts = False
ActiveWindow.Close

Next


End Sub
