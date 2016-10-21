Attribute VB_Name = "Módulo27"
Sub extrai_etp()
Attribute extrai_etp.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro25 Macro
'

'
  
K = 1

For x = 1 To 30

Windows("BH_ESTACOES.xlsX").Activate
Sheets("PLAN2").Select
wth = Range(Cells(1 + x, 1), Cells(1 + x, 1)).Value

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\Interpolado\WTH\" & wth & "_SINTESE.xlsx"

Sheets("BH Sequencial").Select
Range("F19:F1206").Select
Selection.Copy

Windows("BH_ESTACOES.xlsX").Activate
Sheets("PLAN1").Select
Range(Cells(K, 50), Cells(K, 50)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("" & wth & "_SINTESE.xlsx").Activate
Application.DisplayAlerts = False
ActiveWindow.Close
    
K = K + 1188

Next


    

End Sub
