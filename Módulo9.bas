Attribute VB_Name = "Módulo9"
Sub probpro()
Attribute probpro.VB_ProcData.VB_Invoke_Func = " \n14"
'
'
'

'

Windows("Prob_Produ.xlsm").Activate
Sheets("Entrada").Select

dire = ActiveSheet.Dir
cul = Sheets("Entrada").Range("A2").Value
loca = Sheets("Entrada").Range("B2").Value
ciclo = Sheets("Entrada").Range("C2").Value
tole = Sheets("Entrada").Range("D2").Value
datsem = Sheets("Entrada").Range("A4").Value
ciclo = Sheets("Entrada").Range("C2").Value
custo = Sheets("Entrada").Range("D4").Value

Workbooks.Open Filename:="C:\Murilo\MACRO\Prob_Produ\" & cul & "_Dados.xlsx"

Sheets("" & loca & "").Select
Columns("A:G").Select
Selection.AutoFilter
AutoFilter Field:=3, Criteria1:=tole
AutoFilter Field:=4, Criteria1:=ciclo
AutoFilter Field:=5, Criteria1:=datasem
Range("G1").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("Prob_Produ.xlsm").Activate
Range("E4").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("" & cul & "_Dados.xlsx").Activate
Application.DisplayAlerts = False
ActiveWindow.Close
Windows("Prob_Produ.xlsm").Activate
Sheets("Entrada").Select

    
End Sub
