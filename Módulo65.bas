Attribute VB_Name = "Módulo65"
Sub FORMATA_RAD()
Attribute FORMATA_RAD.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro48 Macro
'

'

For x = 293 To 295

Windows("MODELO_JUQUEI.xlsx").Activate
Sheets("ENTRADA").Columns("A:AT").ClearContents

arq = Sheets("CALCULos").Range("M" & x + 1).Value
Workbooks.Open Filename:="C:\Murilo\DOUTORADO\Series Clima\radiacao\estacoes_RadiaçãoSolar\" & arq & ""

Sheets("" & arq & "").Select
Columns("A:AT").Copy

Windows("MODELO_JUQUEI.xlsx").Activate
Sheets("ENTRADA").Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

Calculate
nmes = Sheets("Calculos").Range("F2").Value

Sheets("Calculos").Columns("G:I").ClearContents

K = 6

For y = 1 To nmes

Sheets("ENTRADA").Select
Range(Cells(y + 1, 8), Cells(y + 1, 8 + 30)).Select
Selection.Copy
Sheets("Calculos").Select
Range(Cells(K, 9), Cells(K, 9)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=True

Sheets("ENTRADA").Select
Range(Cells(1, 8), Cells(1, 8 + 30)).Select
Selection.Copy
Sheets("Calculos").Select
Range(Cells(K, 8), Cells(K, 8)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=True

Sheets("ENTRADA").Select
Range(Cells(y + 1, 2), Cells(y + 1, 2)).Select
Selection.Copy
Sheets("Calculos").Select
Range(Cells(K, 7), Cells(K + 30, 7)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

K = K + 31

Next

Calculate

nomenovo = Sheets("ENTRADA").Range("C5").Value
uf = Sheets("ENTRADA").Range("D5").Value

Workbooks.Add
ActiveWorkbook.SaveAs Filename:= _
        "C:\Murilo\DOUTORADO\Series Clima\radiacao\estacoes_RadiaçãoSolar\FORMATADO\" & nomenovo & "_" & uf & ".xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

Windows("MODELO_JUQUEI.xlsx").Activate
Columns("A:B").Select
Selection.Copy
Windows("" & nomenovo & "_" & uf & ".xlsx").Activate
Columns("A:B").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
ActiveWorkbook.Save
Application.DisplayAlerts = False
ActiveWindow.Close

Windows("" & arq & "").Activate
Application.DisplayAlerts = False
ActiveWindow.Close


Next




End Sub

Sub CONSISTENCIA_RAD()
'
' Macro48 Macro
'

'


For x = 1 To 294

Windows("MODELO_CONSISTENCIA.xlsx").Activate
Sheets("Plan1").Columns("A:B").ClearContents

arq = Sheets("lista").Range("A" & x + 2).Value
Workbooks.Open Filename:="C:\Murilo\DOUTORADO\Series Clima\radiacao\estacoes_RadiaçãoSolar\FORMATADO\" & arq & ".xlsx"

Columns("A:B").Copy
Windows("MODELO_CONSISTENCIA.xlsx").Activate
Sheets("Plan1").Select
Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Calculate

Range("E3:P3").Copy
Sheets("lista").Select
Range(Cells(x + 2, 2), Cells(x + 2, 2)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

Windows("" & arq & "").Activate
Application.DisplayAlerts = False
ActiveWindow.Close

Next

End Sub

