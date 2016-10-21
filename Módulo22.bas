Attribute VB_Name = "Módulo22"
Sub Macro16()
Attribute Macro16.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro16 Macro
'

'
    Rows("1:1").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Cells.Select
    Selection.QueryTable.Delete
    Selection.ClearContents
End Sub
