Attribute VB_Name = "Módulo45"
Sub Macro27()
Attribute Macro27.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro27 Macro
'

'

For x = 1 To 11

Windows("controle.xlsx").Activate
Sheets("lista").Select
wth = Sheets("lista").Range("A" & x).Value

Workbooks.Open Filename:="C:\Murilo\DOUTORADO\AGMIP\DataSubmission_Stage1_v2\" & wth & ""
Sheets("Experiment_details").Select
Range("B6:B12").Select
Selection.Copy

Windows("controle.xlsx").Activate
Sheets("lista").Select
Range(Cells(x, 2), Cells(x, 2)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True

Windows("" & wth & "").Activate
Application.DisplayAlerts = False
ActiveWindow.Close

Next

End Sub
