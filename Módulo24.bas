Attribute VB_Name = "Módulo24"
Sub Macro21()
Attribute Macro21.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro21 Macro
'

'
    Range("AM19:BU42").Select
    Selection.Copy
    Range("BP45").Select
    ActiveSheet.Pictures.Paste.Select
    ActiveSheet.Shapes.Range(Array("Picture 11")).Select
    ActiveWindow.SmallScroll Down:=9
    Range("BY41").Select
    ActiveSheet.Shapes.Range(Array("Picture 11")).Select
    Selection.Delete
End Sub
