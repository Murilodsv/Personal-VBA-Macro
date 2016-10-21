Attribute VB_Name = "Módulo69"
Sub Julian_Date()
Attribute Julian_Date.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro55 Macro
'

'

Workbooks.Open Filename:="C:\Users\Murilo\Dropbox\MACRO\Julian_Date.xlsx"



End Sub


Sub PIRA()
'
' Macro55 Macro
'

'

l = 7

For x = 1 To 16

Windows("lista_rob.xlsx").Activate
arquivo = Sheets("Plan1").Range("A" & x).Value
bi = Sheets("Plan1").Range("D" & x).Value

Workbooks.Open Filename:="C:\Murilo\MESTRADO\Dados_met\Dados Diarios Pira\" & arquivo & ""

K = 6

For y = 1 To 12

Windows("lista_rob.xlsx").Activate

If bi = 0 Then
dia = Sheets("Plan1").Range("G" & y).Value
End If
dia = Sheets("Plan1").Range("F" & y).Value

Windows("" & arquivo & "").Activate


Range(Cells(K, 1), Cells(K + dia - 1, 30)).Select
Selection.Copy

Windows("Cópia de PIRA_1917_2015_Total.xlsx").Activate
Range(Cells(l, 1), Cells(l, 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
l = l + dia

K = K + 45

Next

Windows("" & arquivo & "").Activate
Application.DisplayAlerts = False
ActiveWindow.Close


Next


End Sub

