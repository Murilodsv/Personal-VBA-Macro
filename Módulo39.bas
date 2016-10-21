Attribute VB_Name = "Módulo39"
Sub Macro23()
Attribute Macro23.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro23 Macro
'

'
    Range("W21").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[2]:RC[3])"
    Range("W21").Select
    Selection.AutoFill Destination:=Range("W21:W32")
    Range("W21:W32").Select
    Calculate
    Range("W18").Select
End Sub
