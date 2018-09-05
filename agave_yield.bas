Sub extract_agave_prod()
'
' extract_agave_prod Macro
'

'file = D:\Murilo\data\agave\yield_data\MEXICO\agave_tequilana.xslm

nfiles = 15

For x = 3 To nfiles

Workbooks.Open Filename:="D:\Murilo\data\agave\yield_data\MEXICO\Cierre (" & x & ").xls"

Range("K11").Select
ActiveCell.FormulaR1C1 = "=R1C1"
Range("K11").Select
Selection.AutoFill Destination:=Range("K11:K1000")

Range("A11").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("agave_tequilana.xlsm").Activate
l = Sheets("Plan1").Range("Q" & 1).Value

Range(Cells(l, 1), Cells(l, 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("Cierre (" & x & ").xls").Activate
Application.DisplayAlerts = False
ActiveWindow.Close

Next

    
End Sub
