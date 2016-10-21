Attribute VB_Name = "Módulo57"
Sub Macro34()
Attribute Macro34.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro34 Macro
'

'
    Columns("B:B").Select
    Selection.AutoFilter
    Selection.AutoFilter Field:=1, Criteria1:=41751
    
    ActiveSheet.Range("$B$1:$B$1780").AutoFilter Field:=1
    Selection.AutoFilter
    Sheets("Calc").Select
    Range("D7").Select
    Selection.AutoFilter
    Range("H8").Select
    Sheets("INMET").Select
    Range("F18").Select
    ActiveWindow.SmallScroll Down:=1212
    Range("A1222").Select
    ActiveCell.FormulaR1C1 = "=RC[1]*1"
    Range("A1223").Select
    Columns("B:B").Select
    Range("B1213").Activate
    ActiveWindow.SmallScroll Down:=-1434
    ActiveSheet.Range("$B$1:$B$1780").AutoFilter Field:=1, Criteria1:=Array("=" _
        ), Operator:=xlFilterValues, Criteria2:=Array(1, "2/28/2014")
End Sub
Sub Macro35()
Attribute Macro35.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro35 Macro
'

'
    Range("A3").Select
    Selection.End(xlDown).Select
    Range("A19504").Select
End Sub
