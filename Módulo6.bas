Attribute VB_Name = "Módulo6"
Sub ESTATISTICA()
'
' ESTATISTICA MANIPULA A PLANILHA COM VALORES ESTATISTICOS PARA COMPARACAO DE MODELOS
'

'

n_est = Sheets("BASE_ESTAT").Range("R" & 1).Value
n = Sheets("BASE_ESTAT").Range("R" & 2).Value

For x = 1 To n_est

If x = 1 Then

Sheets("BASE_ESTAT").Select
Range("C7:M7").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

Range("C6").Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Copy

Range(Cells(6, 3), Cells(n + 5, 3)).Select
ActiveSheet.Paste

Range("A6:B6").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

Sheets("SAIDA").Select
Range("A3:B3").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

Sheets("ENTRADA").Select
Range(Cells(5, 2), Cells(5, n_est + 1)).Select
Selection.Copy

Sheets("SAIDA").Select
Range("A3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True

End If

Sheets("ENTRADA").Select
Range(Cells(6, 1), Cells(n + 5, 1)).Select
Selection.Copy

Sheets("BASE_ESTAT").Select
Range("A6").Select
ActiveSheet.Paste

Sheets("ENTRADA").Select
Range(Cells(6, x + 1), Cells(n + 5, x + 1)).Select
Selection.Copy

Sheets("BASE_ESTAT").Select
Range("B6").Select
ActiveSheet.Paste

Range("R5").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Sheets("SAIDA").Select
Range(Cells(x + 2, 2), Cells(x + 2, 2)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True

Next
    
End Sub


