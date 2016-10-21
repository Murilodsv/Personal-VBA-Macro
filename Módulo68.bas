Attribute VB_Name = "Módulo68"
Sub Macro53()
Attribute Macro53.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro53 Macro
'

'



For x = 2 To 12


Workbooks.Open Filename:="C:\Users\Murilo\Dropbox\Bills\BB\extrato (" & x & ").csv"

Range(Cells(2, 1), Cells(2, 6)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("Contas_BR.xlsx").Activate
Range("A2").Select
Selection.End(xlDown).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        

Windows("extrato (" & x & ").csv").Activate
Application.DisplayAlerts = False
ActiveWindow.Close


Next

    'ActiveWorkbook.Save
End Sub
Sub Macro54()
Attribute Macro54.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro54 Macro
'

'
    Columns("A:B").Select
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 62
    ActiveWindow.ScrollRow = 92
    ActiveWindow.ScrollRow = 122
    ActiveWindow.ScrollRow = 153
    ActiveWindow.ScrollRow = 183
    ActiveWindow.ScrollRow = 213
    ActiveWindow.ScrollRow = 244
    ActiveWindow.ScrollRow = 274
    ActiveWindow.ScrollRow = 304
    ActiveWindow.ScrollRow = 334
    ActiveWindow.ScrollRow = 395
    ActiveWindow.ScrollRow = 425
    ActiveWindow.ScrollRow = 456
    ActiveWindow.ScrollRow = 486
    ActiveWindow.ScrollRow = 516
    ActiveWindow.ScrollRow = 547
    ActiveWindow.ScrollRow = 577
    ActiveWindow.ScrollRow = 607
    ActiveWindow.ScrollRow = 638
    ActiveWindow.ScrollRow = 668
    ActiveWindow.ScrollRow = 698
    ActiveWindow.ScrollRow = 729
    ActiveWindow.ScrollRow = 759
    ActiveWindow.ScrollRow = 789
    ActiveWindow.ScrollRow = 820
    ActiveWindow.ScrollRow = 789
    ActiveWindow.ScrollRow = 759
    ActiveWindow.ScrollRow = 729
    Selection.Replace What:="#N/D", Replacement:="-999", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub
