Attribute VB_Name = "Módulo2"
Sub AGRUPA_UF()
Attribute AGRUPA_UF.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'

Workbooks.Open Filename:="C:\Murilo\ARCGIS\FGV\MAURO\AREAS.xlsx"


For x = 1 To 21


Windows("AREAS.xlsx").Activate
CSV = Sheets("PLAN2").Range("D" & x).Value
NLINHA = Range("AA1").Value + 1

Workbooks.Open Filename:="C:\Murilo\ARCGIS\FGV\MAURO\CORTE_ZAE\POTENCIAL\ALBERS\AREA\" & CSV & ".csv"

    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("AREAS.xlsx").Activate
    Sheets("PLAN3").Select
    Range(Cells(NLINHA, 1), Cells(NLINHA, 1)).Select
    ActiveSheet.Paste
    
    Windows("" & CSV & ".csv").Activate

    Application.DisplayAlerts = False
    ActiveWindow.Close

Next

End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveWindow.SmallScroll Down:=27
    Range("A94").Select
    ActiveCell.FormulaR1C1 = " "
    Cells.Replace What:=" ", Replacement:="", LookAt:=xlWhole, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End Sub
