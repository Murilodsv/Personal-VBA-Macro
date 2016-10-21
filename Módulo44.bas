Attribute VB_Name = "Módulo44"
Sub criaWTH()
'
' Macro23 Macro
'

'

'Workbooks.Open Filename:="C:\Murilo\MACRO\WTH_DSSAT.xlsx"
Application.Calculation = xlManual
dire = "C:\Murilo\MACRO" 'ThisWorkbook.Path

For x = 1 To 1

Windows("WTH_DSSAT.xlsx").Activate
Sheets("WTH_FINAL").Select

Calculate

Range("A6:A6").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Sheets("LISTA").Select
Range("A1:A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
ActiveSheet.Range("$A$1:$A$12054").RemoveDuplicates Columns:=1, Header:= _
        xlNo
        
Calculate

nano = Sheets("LISTA").Range("C" & 1).Value - 2
wth = Sheets("ENTRADA").Range("B" & 4).Value

For y = 1 To nano

ano = Sheets("LISTA").Range("A" & y + 1).Value

Sheets("EXPORTA").Select
Range("A6:A400").Select
Selection.ClearContents

Sheets("WTH_FINAL").Select
ActiveSheet.Range("$A$5:$A$12058").AutoFilter Field:=1, Criteria1:=ano

Range("U5").Select
Selection.End(xlDown).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Sheets("EXPORTA").Select
Range("A6").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Columns("A:A").Select

Workbooks.Open Filename:="C:\Murilo\MACRO\IMPORTA.xlsx"
Windows("WTH_DSSAT.xlsx").Activate
Sheets("EXPORTA").Select
Columns("A:A").Select
Selection.Copy

Windows("IMPORTA.xlsx").Activate
Columns("A:A").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Application.DisplayAlerts = False
ActiveWorkbook.SaveAs Filename:="" & dire & "\" & wth & "" & ano & "01.WTH", _
       FileFormat:=xlTextPrinter, CreateBackup:=False

Application.DisplayAlerts = False
ActiveWindow.Close
 
 
Next

Sheets("WTH_FINAL").Select
ActiveSheet.Range("$A$5:$A$12058").AutoFilter Field:=1

Next

End Sub



