Attribute VB_Name = "Módulo42"
Sub Macro24()
Attribute Macro24.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro24 Macro
'

'
    Range("E7").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=-51
    Range("E7:E36,I7").Select
    Range("I7").Activate
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("TCH_TUDO.xlsx").Activate
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveWindow.SmallScroll Down:=-30
End Sub
Sub Macro25()
Attribute Macro25.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro25 Macro
'

'
    Range("A7").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=-21
    Range("A7:A36,E7").Select
    Range("E7").Activate
    Range(Selection, Selection.End(xlDown)).Select
    Range("O25").Select
    ActiveWindow.SmallScroll Down:=-45
    Windows("TCH_TUDO.xlsx").Activate
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "SIMULACAO"
    Range("B3").Select
    Columns("B:B").EntireColumn.AutoFit
    Windows("IR011290.xlsx").Activate
    Range("A7:B7").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=-69
    Range("A7:B36,E7").Select
    Range("E7").Activate
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=-15
    Range("A7:B36,E7:E36,I7").Select
    Range("I7").Activate
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("TCH_TUDO.xlsx").Activate
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
