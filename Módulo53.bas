Attribute VB_Name = "Módulo53"
Sub JUNTA()
Attribute JUNTA.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro34 Macro
'

'


Windows("MACRO_JUNTA.xlsm").Activate
arquivo = Range("M2").Value
diretorio = Range("N2").Value
Workbooks.Open Filename:="" & diretorio & "\" & arquivo & ""
linha = 5

For x = 1 To 12

Windows("MACRO_JUNTA.xlsm").Activate
mes = Sheets("JUNTO").Range("O" & x + 1).Value

Windows("" & arquivo & "").Activate
Sheets("" & mes & "").Select

Range("N2").Select
ActiveCell.FormulaR1C1 = "=COUNT(R[3]C[-13]:R[48]C[-13])"

dias = Range("N2").Value

Range("N2").Select
Selection.ClearContents

Range(Cells(5, 1), Cells(4 + dias, 12)).Select
Selection.Copy

Windows("MACRO_JUNTA.xlsm").Activate
Sheets("JUNTO").Select

Range(Cells(linha, 1), Cells(linha, 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

linha = linha + dias

Next


End Sub
