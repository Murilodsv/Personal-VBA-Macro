Attribute VB_Name = "Módulo63"
Sub Macro46()
Attribute Macro46.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro46 Macro
'

'
    Range("Q1:R1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.AutoFilter
    ActiveSheet.Range("$Q$1:$R$17").AutoFilter Field:=1, Criteria1:=">1.5", _
        Operator:=xlAnd
    ActiveSheet.Range("$Q$1:$R$17").AutoFilter Field:=2, Criteria1:=">20", _
        Operator:=xlAnd
    Range("E1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("SAIDA").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SAIDA").Sort.SortFields.Add Key:=Range("E1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("SAIDA").Sort
        .SetRange Range("A3:R17")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
End Sub
