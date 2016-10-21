Attribute VB_Name = "Módulo25"
Sub Macro22()
Attribute Macro22.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro22 Macro
'

'
    Columns("D:G").Select
    Selection.Delete Shift:=xlToLeft
    Windows("OUTPUTS_DSSAT_IMPORTA.xlsx").Activate
    Sheets("OPG").Select
    ActiveSheet.Range("$A$1:$C$92480").AutoFilter Field:=1, Criteria1:="ALMA"
    Range("A1").Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Application.CutCopyMode = False
End Sub
