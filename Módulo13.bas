Attribute VB_Name = "Módulo13"
Sub Merge()
Attribute Merge.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Merge
'

'
   
Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\estacoes_selecao.xlsx"
Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\MERGE_SERIE.xlsx"

For x = 1 To 30

Windows("estacoes_selecao.xlsx").Activate

INMET = Sheets("estacoes_selecao").Range("D" & x + 1).Value
ANA = Sheets("estacoes_selecao").Range("AD" & x + 1).Value

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\selecao\" & INMET & ".xlsx"

Columns("A:I").Select
Selection.Copy

Windows("MERGE_SERIE.xlsx").Activate
Sheets("original").Select
Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Selection.Replace What:="", Replacement:="-99", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Windows("" & INMET & ".xlsx").Activate
Application.DisplayAlerts = False
ActiveWindow.Close

Workbooks.Open Filename:="C:\Murilo\MESTRADO\ANA\" & ANA & "_formatado.xlsx"

Columns("A:E").Select
Selection.Copy

Windows("MERGE_SERIE.xlsx").Activate
Sheets("proxima").Select
Range("B1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Selection.Replace What:="", Replacement:="-99", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Windows("" & ANA & "_formatado.xlsx").Activate
Application.DisplayAlerts = False
ActiveWindow.Close

Windows("MERGE_SERIE.xlsx").Activate
Sheets("original").Select
Range("A1:B4").Select
Selection.Copy

Sheets("seriepadrao").Select
Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Columns("A:E").Select
Selection.Copy

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\Modelo_estacao.xlsx"

Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

ActiveWorkbook.SaveAs Filename:= _
        "C:\Murilo\MESTRADO\INMET\selecao\Merge_ANA\" & INMET & "_merge.xls" _
        , FileFormat:=xlExcel8, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
Application.DisplayAlerts = False
ActiveWindow.Close

Windows("MERGE_SERIE.xlsx").Activate
Sheets("original").Select
Columns("A:O").Select
Selection.ClearContents

Sheets("proxima").Select
Columns("A:O").Select
Selection.ClearContents

Next

   

End Sub
