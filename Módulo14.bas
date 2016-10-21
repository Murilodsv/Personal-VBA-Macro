Attribute VB_Name = "Módulo14"
Sub latlong_ana()
Attribute latlong_ana.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro17 Macro
'

'
    
Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\estacoes_selecao.xlsx"

For x = 1 To 30

Windows("estacoes_selecao.xlsx").Activate
ANA = Sheets("estacoes_selecao").Range("AD" & x + 1).Value

Workbooks.Open Filename:="C:\Murilo\MESTRADO\ANA\" & ANA & "_formatado.xlsx"

Range("A3:B3").Select
Selection.Copy

Windows("estacoes_selecao.xlsx").Activate
Range(Cells(1 + x, 31), Cells(1 + x, 31)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("" & ANA & "_formatado.xlsx").Activate
Application.DisplayAlerts = False
ActiveWindow.Close

Next
   
    
End Sub

Sub Angstron()
'
' Macro17 Macro
'

'
    
Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\estacoes_selecao.xlsx"
Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\RADIACAO.xlsx"

For x = 1 To 30

Windows("estacoes_selecao.xlsx").Activate
INMET = Sheets("estacoes_selecao").Range("D" & x + 1).Value

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\" & INMET & "_merge.xls"

Columns("A:E").Select
Selection.Copy

Windows("RADIACAO.xlsx").Activate
Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("estacoes_selecao.xlsx").Activate
Range(Cells(1 + x, 33), Cells(1 + x, 34)).Select
Selection.Copy

Windows("RADIACAO.xlsx").Activate
Range("G2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("estacoes_selecao.xlsx").Activate
Range(Cells(1 + x, 35), Cells(1 + x, 36)).Select
Selection.Copy

Windows("RADIACAO.xlsx").Activate
Range("G3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

ang = Range("G3").Value

ActiveWorkbook.Save

Range("A7").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("" & INMET & "_merge.xls").Activate
Range("A7").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
If ang > 0.01 Then

    Columns("F:K").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.ClearContents
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "SRAD"

Else

    Columns("F:L").Select
    Selection.Delete Shift:=xlToLeft
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "SRAD"

End If

ActiveWorkbook.SaveAs Filename:= _
        "C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Radiacao\" & INMET & "_merge_Rad.xls" _
        , FileFormat:=xlExcel8, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
Application.DisplayAlerts = False
ActiveWindow.Close


Windows("RADIACAO.xlsx").Activate
Columns("A:E").Select
Selection.ClearContents

Next
   
    
End Sub


