Attribute VB_Name = "Módulo41"
Sub BD_NFs()
'
'
'

'

    For X = 1 To 523
    
    Workbooks.Open Filename:="C:\Users\Murilo\Dropbox\MURILO\Mini_Curso_Excel\Banco de Dados\MODELO_NF.xls"

    Windows("Banco de NF.xlsx").Activate
    Sheets("Banco_de_NF").Select
    
    NOTA = Sheets("Banco_de_NF").Range("A" & X + 1).Value
    
    Range(Cells(X + 1, 1), Cells(X + 1, 1)).Select
    Selection.Copy
    
    Windows("MODELO_NF.xls").Activate
    Range("AA5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Windows("Banco de NF.xlsx").Activate
    Sheets("Banco_de_NF").Select
    
    Range(Cells(X + 1, 2), Cells(X + 1, 2)).Select
    Selection.Copy
    
    Windows("MODELO_NF.xls").Activate
    Range("AA23").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    Windows("Banco de NF.xlsx").Activate
    Sheets("Banco_de_NF").Select
    
    Range(Cells(X + 1, 4), Cells(X + 1, 4)).Select
    Selection.Copy
    
    Windows("MODELO_NF.xls").Activate
    Range("J23").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    Windows("Banco de NF.xlsx").Activate
    Sheets("Banco_de_NF").Select
    
    Range(Cells(X + 1, 5), Cells(X + 1, 5)).Select
    Selection.Copy
    
    Windows("MODELO_NF.xls").Activate
    Range("Q58").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    Windows("Banco de NF.xlsx").Activate
    Sheets("Banco_de_NF").Select
    
    Range(Cells(X + 1, 6), Cells(X + 1, 6)).Select
    Selection.Copy
    
    Windows("MODELO_NF.xls").Activate
    Range("V27").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    Windows("Banco de NF.xlsx").Activate
    Sheets("Banco_de_NF").Select
    
    Range(Cells(X + 1, 7), Cells(X + 1, 7)).Select
    Selection.Copy
    
    Windows("MODELO_NF.xls").Activate
    Range("V23").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    Windows("Banco de NF.xlsx").Activate
    Sheets("Banco_de_NF").Select
    
    Range(Cells(X + 1, 8), Cells(X + 1, 8)).Select
    Selection.Copy
    
    Windows("MODELO_NF.xls").Activate
    Range("R27").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Windows("Banco de NF.xlsx").Activate
    Sheets("Banco_de_NF").Select
    
    Range(Cells(X + 1, 9), Cells(X + 1, 9)).Select
    Selection.Copy
    
    Windows("MODELO_NF.xls").Activate
    Range("J25").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Windows("Banco de NF.xlsx").Activate
    Sheets("Banco_de_NF").Select
    
    Range(Cells(X + 1, 10), Cells(X + 1, 10)).Select
    Selection.Copy
    
    Windows("MODELO_NF.xls").Activate
    Range("J27").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Windows("Banco de NF.xlsx").Activate
    Sheets("Banco_de_NF").Select
    
    Range(Cells(X + 1, 11), Cells(X + 1, 11)).Select
    Selection.Copy
    
    Windows("MODELO_NF.xls").Activate
    Range("U27").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Windows("Banco de NF.xlsx").Activate
    Sheets("Banco_de_NF").Select
    
    Range(Cells(X + 1, 12), Cells(X + 1, 12)).Select
    Selection.Copy
    
    Windows("MODELO_NF.xls").Activate
    Range("W25").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Calculate
    
    Columns("A:AC").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Columns("AC:AZ").Select
    Selection.ClearContents
    
    Range("A1").Select
    
    ActiveWorkbook.SaveAs Filename:= _
        "C:\Users\Murilo\Dropbox\MURILO\Mini_Curso_Excel\Banco de Dados\" & X & ".xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
        Application.DisplayAlerts = False
        ActiveWindow.Close

       
    
    Next
    
End Sub
