Attribute VB_Name = "M�dulo40"
Sub organiza_estacoes_ana()
'
'
'

'



For x = 1 To 1

Workbooks.Open Filename:="C:\Murilo\MACRO\DSSAT\OUTPUT\an�lise\final\analise.xlsx"

BH = Sheets("lista").Range("A" & x + 1).Value
BHP = Sheets("lista").Range("B" & x + 1).Value

Workbooks.Open Filename:="C:\Murilo\MACRO\DSSAT\OUTPUT\an�lise\final\" & BH & ".xlsx"

Sheets("OPG").Select
Cells.Select
Selection.Copy
Windows("analise.xlsx").Activate
Sheets("OPG").Select
Cells.Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

Windows("" & BH & ".xlsx").Activate
Sheets("OSW").Select
Cells.Select
Selection.Copy
Windows("analise.xlsx").Activate
Sheets("OSW").Select
Cells.Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Windows("" & BH & ".xlsx").Activate
Sheets("OEB").Select
Columns("A:Z").Select
Selection.Copy
Windows("analise.xlsx").Activate
Sheets("OEB").Select
Columns("A:Z").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

Workbooks.Open Filename:="C:\Murilo\MACRO\DSSAT\OUTPUT\an�lise\final\" & BHP & ".xlsx"
Sheets("OPG").Select
Cells.Select
Selection.Copy
Windows("analise.xlsx").Activate
Sheets("OPG_P").Select
Cells.Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

Calculate

ActiveWorkbook.SaveAs Filename:="C:\Murilo\MACRO\DSSAT\OUTPUT\an�lise\final\" & BH & "_analise.xlsx", _
FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
ActiveWindow.Close

Windows("" & BH & ".xlsx").Activate
Application.DisplayAlerts = False
ActiveWindow.Close

Windows("" & BHP & ".xlsx").Activate
ActiveWindow.Close

Next

End Sub

