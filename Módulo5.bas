Attribute VB_Name = "Módulo5"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveSheet.ChartObjects("Chart 3").Activate
    ActiveWindow.SmallScroll Down:=9
    ActiveSheet.Shapes.Range(Array("Chart 3", "Chart 4")).Select
    Selection.Copy
    Windows("Cola_Grafico.xlsx").Activate
    Sheets("10").Select
    Range("B3").Select
    ActiveSheet.Pictures.Paste.Select
End Sub
