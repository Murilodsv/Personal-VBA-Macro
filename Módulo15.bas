Attribute VB_Name = "Módulo15"
Sub Interpolador_TEMP()
Attribute Interpolador_TEMP.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro16 Macro
'

'
Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\estacoes_selecao.xlsx"
Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\Interpolado\Interpolador.xls"

For x = 1 To 30

Windows("estacoes_selecao.xlsx").Activate
INMET = Sheets("estacoes_selecao").Range("D" & x + 1).Value

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\" & INMET & "_merge_Rad.xls"

Range("A6").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("Interpolador.xls").Activate
Range("A6").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

N_LINHA = Sheets("Plan1").Range("G" & 2).Value - 1

Range(Cells(7, 9), Cells(7 + N_LINHA, 9)).Select
Selection.Copy

Windows("" & INMET & "_merge_Rad.xls").Activate
Range("C7").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Windows("Interpolador.xls").Activate
Range(Cells(7, 12), Cells(7 + N_LINHA, 12)).Select
Selection.Copy

Windows("" & INMET & "_merge_Rad.xls").Activate
Range("D7").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False


ActiveWorkbook.SaveAs Filename:= _
        "C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\Interpolado\" & INMET & "_merge_Rad_int.xls" _
        , FileFormat:=xlExcel8, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
Application.DisplayAlerts = False
ActiveWindow.Close

Windows("Interpolador.xls").Activate
Columns("A:F").Select
Selection.ClearContents

Next


    
End Sub
