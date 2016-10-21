Attribute VB_Name = "Módulo11"
Sub Macro15()
Attribute Macro15.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro15 Macro
'

'

ano = 0

For x = 1 To 10

Workbooks.Open Filename:="C:\Murilo\MESTRADO\shape\CanaSat_Tabelas\GO_20" & x + 2 & "-" & x + 3 & ".xls"

Range(Cells(7, 3), Cells(7, 7)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
    
Windows("Cansat_2003_2012.xlsx").Activate
Sheets("GO_2003-2012").Select

Range(Cells(7, 3 + ano), Cells(7, 3 + ano)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("GO_20" & x + 2 & "-" & x + 3 & ".xls").Activate
Application.DisplayAlerts = False
ActiveWindow.Close


Workbooks.Open Filename:="C:\Murilo\MESTRADO\shape\CanaSat_Tabelas\MG_20" & x + 2 & "-" & x + 3 & ".xls"

Range(Cells(7, 3), Cells(7, 7)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
    
Windows("Cansat_2003_2012.xlsx").Activate
Sheets("MG_2003-2012").Select

Range(Cells(7, 3 + ano), Cells(7, 3 + ano)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("MG_20" & x + 2 & "-" & x + 3 & ".xls").Activate
Application.DisplayAlerts = False
ActiveWindow.Close


Workbooks.Open Filename:="C:\Murilo\MESTRADO\shape\CanaSat_Tabelas\SP_20" & x + 2 & "-" & x + 3 & ".xls"

Range(Cells(7, 3), Cells(7, 7)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
    
Windows("Cansat_2003_2012.xlsx").Activate
Sheets("SP_2003-2012").Select

Range(Cells(7, 3 + ano), Cells(7, 3 + ano)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("SP_20" & x + 2 & "-" & x + 3 & ".xls").Activate
Application.DisplayAlerts = False
ActiveWindow.Close

Workbooks.Open Filename:="C:\Murilo\MESTRADO\shape\CanaSat_Tabelas\MS_20" & x + 2 & "-" & x + 3 & ".xls"

Range(Cells(7, 3), Cells(7, 7)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
    
Windows("Cansat_2003_2012.xlsx").Activate
Sheets("MS_2003-2012").Select

Range(Cells(7, 3 + ano), Cells(7, 3 + ano)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("MS_20" & x + 2 & "-" & x + 3 & ".xls").Activate
Application.DisplayAlerts = False
ActiveWindow.Close


Workbooks.Open Filename:="C:\Murilo\MESTRADO\shape\CanaSat_Tabelas\PR_20" & x + 2 & "-" & x + 3 & ".xls"

Range(Cells(7, 3), Cells(7, 7)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
    
Windows("Cansat_2003_2012.xlsx").Activate
Sheets("PR_2003-2012").Select

Range(Cells(7, 3 + ano), Cells(7, 3 + ano)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("PR_20" & x + 2 & "-" & x + 3 & ".xls").Activate
Application.DisplayAlerts = False
ActiveWindow.Close
    
    
Workbooks.Open Filename:="C:\Murilo\MESTRADO\shape\CanaSat_Tabelas\MT_20" & x + 2 & "-" & x + 3 & ".xls"

Range(Cells(7, 3), Cells(7, 7)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
    
Windows("Cansat_2003_2012.xlsx").Activate
Sheets("MT_2003-2012").Select

Range(Cells(7, 3 + ano), Cells(7, 3 + ano)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("MT_20" & x + 2 & "-" & x + 3 & ".xls").Activate
Application.DisplayAlerts = False
ActiveWindow.Close

ano = ano + 5

Next

End Sub
