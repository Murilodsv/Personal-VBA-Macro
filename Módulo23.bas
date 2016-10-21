Attribute VB_Name = "Módulo23"
Sub Macro17()
Attribute Macro17.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro17 Macro
'

'
    ActiveSheet.ChartObjects("Chart 3").Activate
    ActiveChart.ChartArea.Copy
    Range("S8").Select
    ActiveSheet.ChartObjects("Chart 3").Activate
    ActiveSheet.Shapes.Range(Array("Chart 3", "Chart 4", "Chart 3")).Select
    Selection.Copy
    Range("S8").Select
    ActiveSheet.ChartObjects("Chart 3").Activate
    ActiveChart.ChartArea.Copy
    Range("S9").Select
End Sub
Sub Macro19()
Attribute Macro19.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro19 Macro
'

'
    ActiveWindow.SmallScroll ToRight:=-5
    ActiveSheet.ChartObjects("Chart 3").Activate
    ActiveSheet.Shapes.Range(Array("Chart 3", "Chart 3")).Select
    ActiveSheet.Shapes.Range(Array("Chart 3", "Chart 3", "Chart 4")).Select
    Range("U8").Select
    ActiveSheet.Paste
    Selection.Delete
End Sub
Sub Macro20()
Attribute Macro20.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro20 Macro
'

'
    ActiveSheet.ChartObjects("Chart 3").Activate
    ActiveSheet.Shapes("Chart 3").Name = "Chart 5"
    Range("S7").Select
End Sub
