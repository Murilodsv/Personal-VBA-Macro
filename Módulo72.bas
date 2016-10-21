Attribute VB_Name = "Módulo72"
Sub graficos_bowen()
'
' graficos_bowen Macro
'

'

For x = 1 To 144

Windows("Modelo_Graficos_SP.xlsx").Activate
filtro = Sheets("Lista Filtro").Range("A" & x + 1).Value
ET = Sheets("Lista Filtro").Range("B" & x + 1).Value


Sheets("Plan3").Select

    If x = 1 Then
    l = 1
    Range("X7:X26834").Select
    Selection.AutoFilter
    End If
    ActiveSheet.Range("$X$7:$X$26834").AutoFilter Field:=1, Criteria1:= _
        "=" & filtro & "", Operator:=xlAnd
        
    Range("A7:AK7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    
    Sheets("Plan1").Select
    Range("B7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Calculate

    Sheets("Plan2").Select
    Range(Cells(l, 1), Cells(l, 1)).Select
    ActiveCell.FormulaR1C1 = filtro
    Range(Cells(l, 2), Cells(l, 2)).Select
    ActiveCell.FormulaR1C1 = ET
    
    Sheets("Plan1").Select
    ActiveSheet.ChartObjects("Gráfico 2").Activate
    ActiveSheet.Shapes.Range(Array("Chart 2", "Chart 3")).Select
    Selection.Copy
    Sheets("Plan2").Select
    Range(Cells(l + 3, 1), Cells(l + 3, 1)).Select
    ActiveSheet.Pictures.Paste.Select
    
    l = l + 25
    
Next

End Sub


