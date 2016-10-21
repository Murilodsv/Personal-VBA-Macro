Attribute VB_Name = "Módulo66"
Sub TEMPORARIAS()
Attribute TEMPORARIAS.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro49 Macro
'

'
K = 2
col = 4

For x = 1 To 70

Sheets("2001_2002").Select

Range(Cells(6, K), Cells(5501, K + 1)).Select
Selection.Copy

Sheets("2001_2012").Select
Range(Cells(6, col), Cells(5559, col)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
  
Sheets("2003_2004").Select

Range(Cells(6, K), Cells(5501, K + 1)).Select
Selection.Copy

Sheets("2001_2012").Select
col = col + 2
Range(Cells(6, col), Cells(5501, col)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Sheets("2005_2006").Select

Range(Cells(6, K), Cells(5501, K + 1)).Select
Selection.Copy

Sheets("2001_2012").Select
col = col + 2
Range(Cells(6, col), Cells(5501, col)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Sheets("2007_2008").Select

Range(Cells(6, K), Cells(5501, K + 1)).Select
Selection.Copy

Sheets("2001_2012").Select
col = col + 2
Range(Cells(6, col), Cells(5501, col)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Sheets("2009_2010").Select

Range(Cells(6, K), Cells(5501, K + 1)).Select
Selection.Copy

Sheets("2001_2012").Select
col = col + 2
Range(Cells(6, col), Cells(5501, col)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Sheets("2011_2012").Select

Range(Cells(6, K), Cells(5501, K + 1)).Select
Selection.Copy

Sheets("2001_2012").Select
col = col + 2
Range(Cells(6, col), Cells(5501, col)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
col = col + 2

K = K + 2

Next

    
End Sub
Sub Macro50()
Attribute Macro50.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro50 Macro
'

'
    Sheets.Add After:=Sheets(Sheets.Count)
    Range("K20").Select
    Sheets("Plan18").Select
    Sheets("Plan18").Name = "abc"
    Range("K22").Select
End Sub
Sub Macro51()
Attribute Macro51.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro51 Macro
'

'

For x = 1 To 301

Windows("Lista_Estacoes.xlsx").Activate

arq = Sheets("lista").Range("A" & x + 1).Value

Workbooks.Open Filename:="C:\Murilo\DOUTORADO\Series Clima\radiacao\estacoes_RadiaçãoSolar\" & arq & ""

Sheets("" & arq & "").Select
Range("A5:G5").Select
Selection.Copy

Windows("Lista_Estacoes.xlsx").Activate
Range(Cells(x + 1, 2), Cells(x + 1, 2)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("" & arq & "").Activate

Application.DisplayAlerts = False
ActiveWindow.Close

    
Next

End Sub
