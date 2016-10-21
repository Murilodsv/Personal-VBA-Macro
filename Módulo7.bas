Attribute VB_Name = "Módulo7"
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'

    Windows("sensibilidade_Century.xlsx").Activate
    Sheets("PP").Select
    
    Range("AJ4:AJ1203").Select
    Selection.Copy
    
    Windows("Estatistica1.1.xlsm").Activate
    Sheets("ENTRADA").Select
    Range("A6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    COLUNA = 6
    K = 2
    For x = 1 To 4
        
    Windows("sensibilidade_Century.xlsx").Activate
    Sheets("PP").Select
    
    Range(Cells(4, COLUNA), Cells(1203, COLUNA)).Select
    Selection.Copy
        
    Windows("Estatistica1.1.xlsm").Activate
    Sheets("ENTRADA").Select
    Range(Cells(6, K), Cells(6, K)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    K = K + 1
    COLUNA = COLUNA + 8
    
    Next
    
    COLUNA = 6
    
    For y = 1 To 4
        
    Windows("sensibilidade_Century.xlsx").Activate
    Sheets("T").Select
    
    Range(Cells(4, COLUNA), Cells(1203, COLUNA)).Select
    Selection.Copy
        
    Windows("Estatistica1.1.xlsm").Activate
    Sheets("ENTRADA").Select
    Range(Cells(6, K), Cells(6, K)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    K = K + 1
    COLUNA = COLUNA + 8
    
    Next
    
    
    COLUNA = 6
    
    For Z = 1 To 4
        
    Windows("sensibilidade_Century.xlsx").Activate
    Sheets("Bulk").Select
    
    Range(Cells(4, COLUNA), Cells(1203, COLUNA)).Select
    Selection.Copy
        
    Windows("Estatistica1.1.xlsm").Activate
    Sheets("ENTRADA").Select
    Range(Cells(6, K), Cells(6, K)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    K = K + 1
    COLUNA = COLUNA + 8
    
    Next
    
    COLUNA = 6
    For J = 1 To 4
        
    Windows("sensibilidade_Century.xlsx").Activate
    Sheets("PH").Select
    
    Range(Cells(4, COLUNA), Cells(1203, COLUNA)).Select
    Selection.Copy
        
    Windows("Estatistica1.1.xlsm").Activate
    Sheets("ENTRADA").Select
    Range(Cells(6, K), Cells(6, K)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    K = K + 1
    COLUNA = COLUNA + 8
    
    Next
    
    COLUNA = 6
    For J = 1 To 4
        
    Windows("sensibilidade_Century.xlsx").Activate
    Sheets("CLAY").Select
    
    Range(Cells(4, COLUNA), Cells(1203, COLUNA)).Select
    Selection.Copy
        
    Windows("Estatistica1.1.xlsm").Activate
    Sheets("ENTRADA").Select
    Range(Cells(6, K), Cells(6, K)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    K = K + 1
    COLUNA = COLUNA + 8
    
    Next
    
    COLUNA = 6
    For J = 1 To 4
        
    Windows("sensibilidade_Century.xlsx").Activate
    Sheets("Long").Select
    
    Range(Cells(4, COLUNA), Cells(1203, COLUNA)).Select
    Selection.Copy
        
    Windows("Estatistica1.1.xlsm").Activate
    Sheets("ENTRADA").Select
    Range(Cells(6, K), Cells(6, K)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    K = K + 1
    COLUNA = COLUNA + 8
    
    Next
    

End Sub
