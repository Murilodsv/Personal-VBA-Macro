Attribute VB_Name = "Módulo60"
Sub Macro38()
Attribute Macro38.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro38 Macro
'

'
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "=COUNT(COMPARA!C1)"
    Range("R3").Select
    ActiveCell.FormulaR1C1 = "=COUNT(COMPARA!C1)"
    Range("R4").Select
    ActiveCell.FormulaR1C1 = "CONT.NÚM(COMPARA!A:A)"
    Range("R4").Select
    ActiveCell.FormulaR1C1 = "=COUNT(COMPARA!C[-17])"
    Range("R5").Select
End Sub
