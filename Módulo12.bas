Attribute VB_Name = "Módulo12"
Sub organiza_estacoes_ana()
Attribute organiza_estacoes_ana.VB_ProcData.VB_Invoke_Func = " \n14"
'
'
'

'

For x = 1 To 2

Workbooks.Open Filename:="C:\Murilo\MESTRADO\ANA\modelo_prec.xlsx"
Windows("modelo_prec.xlsx").Activate

ANA = Sheets("lista").Range("A" & x).Value

Workbooks.Open Filename:="C:\Murilo\MESTRADO\ANA\" & ANA & ".xlsx"

Range(Cells(5, 5), Cells(5, 36)).Select

Cells.Replace What:="Dia", Replacement:="", LookAt:=xlPart, _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False

Range("G1").Select
ActiveCell.FormulaR1C1 = "=COUNTIF(C[-6],1)"
n = Range("G1").Value
dia1 = 6

For y = 1 To n

Windows("" & ANA & "").Activate
mes = Range(Cells(5 + y, 3), Cells(5 + y, 3)).Value
ano = Range(Cells(5 + y, 2), Cells(5 + y, 2)).Value

Range("F4").Select
ActiveCell.FormulaR1C1 = "=(R[" & y + 1 & "]C[-4]/4)-TRUNC(R[" & y + 1 & "]C[-4]/4)"

bis = Range("F4").Value

If bis = 0 Then

Windows("modelo_prec.xlsx").Activate
ndia = Sheets("lista").Range("E" & mes).Value

Else

Windows("modelo_prec.xlsx").Activate
ndia = Sheets("lista").Range("D" & mes).Value

End If

Windows("" & ANA & "").Activate
Range(Cells(5 + y, 2), Cells(5 + y, 2)).Select
Selection.Copy

Windows("modelo_prec.xlsx").Activate
Sheets("plan1").Select
Range(Cells(dia1, 3), Cells(dia1 + ndia - 1, 3)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
Windows("" & ANA & "").Activate
Range(Cells(5 + y, 5), Cells(5 + y, 4 + ndia)).Select
Selection.Copy

Windows("modelo_prec.xlsx").Activate
Sheets("plan1").Select
Range(Cells(dia1, 4), Cells(dia1 + ndia, 4)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
        
Range(Cells(dia1, 2), Cells(dia1 + ndia, 2)).Select
ActiveCell.FormulaR1C1 = mes

Range(Cells(dia1, 2), Cells(dia1, 2)).Select
Selection.Copy

Range(Cells(dia1, 2), Cells(dia1 + ndia - 1, 2)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("" & ANA & "").Activate
Range(Cells(5, 5), Cells(5, 4 + ndia)).Select
Selection.Copy

Windows("modelo_prec.xlsx").Activate
Range(Cells(dia1, 1), Cells(dia1 + ndia - 1, 1)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True


dia1 = dia1 + ndia


Next

Windows("" & ANA & "").Activate
Range(Cells(1, 1), Cells(3, 15)).Select
Selection.Copy

Windows("modelo_prec.xlsx").Activate
Range(Cells(1, 1), Cells(3, 15)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

ActiveWorkbook.SaveAs Filename:="C:\Murilo\MESTRADO\ANA\" & ANA & "_formatado.xlsx", _
FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
ActiveWindow.Close

Windows("" & ANA & "").Activate
Application.DisplayAlerts = False
ActiveWindow.Close


Next

End Sub


Sub CONTADOR_SERIE()
Attribute CONTADOR_SERIE.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro21 Macro
'

'

For x = 1 To 290

Windows("CONTADOR_SERIE.xlsx").Activate
nome = Sheets("Plan2").Range("E" & x + 3).Value

Workbooks.Open Filename:="C:\Murilo\MESTRADO\INMET\Diarios_org\" & nome & ".xlsx"

Range("A6").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
    
Windows("CONTADOR_SERIE.xlsx").Activate
Sheets("plan1").Select
Range("C6").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
Range("H2:J2").Select
Selection.Copy
Sheets("plan2").Select
Range(Cells(x + 3, 29), Cells(x + 3, 29)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Sheets("plan1").Select
Range("H3:J3").Select
Selection.Copy
Sheets("plan2").Select
Range(Cells(x + 3, 32), Cells(x + 3, 32)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Sheets("plan1").Select
Range("C6").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

Windows("" & nome & ".xlsx").Activate
Application.DisplayAlerts = False
ActiveWindow.Close


Next

    
End Sub
