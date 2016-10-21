Attribute VB_Name = "Módulo36"
Sub AGRUPA_TOTAIS()
Attribute AGRUPA_TOTAIS.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro18 Macro
'

'

Workbooks.Open Filename:="C:\Murilo\MACRO\DSSAT\BATCH_DSSAT_TCH\TOTAIS.xlsx"

K = 1

For x = 1 To 12

Windows("TOTAIS.xlsx").Activate
totais = Sheets("LISTA").Range("D" & K + 1).Value

Workbooks.Open Filename:="C:\Murilo\MACRO\DSSAT\BATCH_DSSAT_TCH\OUTPUT\" & totais & ".xlsx"

Sheets("TOTAL").Select
Range("K5:K964").Select
Selection.Copy

Windows("TOTAIS.xlsx").Activate
Sheets("Solo1").Select
Range(Cells(4, 13 + x), Cells(4, 13 + x)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & totais & ".xlsx").Activate
Application.DisplayAlerts = False
ActiveWindow.Close

K = K + 1

Next


For y = 1 To 3

For Z = 1 To 12

Windows("TOTAIS.xlsx").Activate
totais = Sheets("LISTA").Range("D" & K + 1).Value

Workbooks.Open Filename:="C:\Murilo\MACRO\DSSAT\BATCH_DSSAT_TCH\OUTPUT\" & totais & ".xlsx"

Sheets("TOTAL").Select
Range("K5:K964").Select
Selection.Copy

Windows("TOTAIS.xlsx").Activate
Sheets("Solo" & y & "").Select
Range(Cells(4, 1 + Z), Cells(4, 1 + Z)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("" & totais & ".xlsx").Activate
Sheets("MEDIA_TOTAL").Select
Range("AP5:AP34").Select
Selection.Copy

Windows("TOTAIS.xlsx").Activate
Sheets("Solo" & y & "").Select
Range(Cells(4, 76 + Z), Cells(4, 76 + Z)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("" & totais & ".xlsx").Activate
Application.DisplayAlerts = False
ActiveWindow.Close

K = K + 1

Next

Next

Calculate

End Sub
