Attribute VB_Name = "Módulo29"
Sub macro_etp()
Attribute macro_etp.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro23 Macro
'

'

linha = 2
For x = 1 To 30

Windows("OUTPUTS_DSSAT_IMPORTA.xlsm").Activate
Sheets("ETP").Select
wth = Sheets("Lista").Range("A" & x + 1).Value

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\Interpolado\WTH\" & wth & "_SINTESE.xlsx"

Sheets("BH Sequencial").Select
Range("N19:O1206").Select
Selection.Copy

Windows("OUTPUTS_DSSAT_IMPORTA.xlsm").Activate
Sheets("ETP").Select
Range(Cells(linha, 8), Cells(linha, 8)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("" & wth & "_SINTESE.xlsx").Activate
Application.DisplayAlerts = False
ActiveWindow.Close

linha = linha + 1188


Next

End Sub
