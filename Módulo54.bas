Attribute VB_Name = "Módulo54"
Sub Simulacoes_soja()
Attribute Simulacoes_soja.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro35 Macro
'

'

dire = ThisWorkbook.Path
nome = ThisWorkbook.Name

loop_esta = Sheets("Entrada").UsedRange.Rows.Count - 1

For x = 1 To loop_esta

esta = Sheets("Entrada").Range("A" & x + 1).Value
arqu = Sheets("Entrada").Range("C" & x + 1).Value
dsem = Sheets("Entrada").Range("D" & x + 1).Value
numd = Sheets("Entrada").Range("E" & x + 1).Value
ccul = Sheets("Entrada").Range("F" & x + 1).Value
icol = Sheets("Entrada").Range("G" & x + 1).Value
upor = Sheets("Entrada").Range("H" & x + 1).Value

Workbooks.Open Filename:="" & dire & "\" & arqu & ".csv"

Selection.find(What:=dsem - 60, After:=ActiveCell, LookIn:=xlFormulas _
        , LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select

Range(ActiveCell, ActiveCell.Offset(ccul + 60, 6)).Copy

    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=COUNT(R[3]C[-13]:R[48]C[-13])"
    Range("N3").Select
    
    
    
Next


End Sub
