Attribute VB_Name = "Módulo67"
Sub Macro49()
Attribute Macro49.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro49 Macro
'

'
    Workbooks.Add
    Windows("MODELO_JUQUEI.xlsx").Activate
    Columns("A:B").Select
    Selection.Copy
    Windows("Pasta1").Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("C7").Select
End Sub
Sub Macro52()
Attribute Macro52.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro52 Macro
'

'
    ActiveWorkbook.SaveAs Filename:= _
        "C:\Murilo\DOUTORADO\Series Clima\radiacao\estacoes_RadiaçãoSolar\fwefwef.xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Range("O17").Select
    ActiveWindow.Close
End Sub
