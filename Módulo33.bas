Attribute VB_Name = "Módulo33"
Sub Convert_SWAP()
Attribute Convert_SWAP.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Format weather data to SWAP weather files format
'

'

Dim extensao As String

For x = 1 To 9

Windows("Convert_Weather.xlsx").Activate

Call limpa_plan

Sheets("Lista").Select
ABA = Range("C" & x + 1).Value
SIG = Range("B" & x + 1).Value

Range(Cells(x + 1, 17), Cells(x + 1, 17)).Select
Selection.Copy
Sheets("ENTRADA").Select
Range("A7:A20000").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

Windows("Weather_Soil_Data.xlsx").Activate
Sheets("" & ABA & "").Select
Range("A1:I1").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("Convert_Weather.xlsx").Activate
Sheets("ENTRADA").Select
Range("S6").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Calculate

min_ano = Range("M" & 1).Value
max_ano = Range("N" & 1).Value
n_ano = Range("O" & 1).Value

For y = 1 To n_ano

Sheets("PROCESS").Select

Range("D5").Select
Range(Selection, Selection.End(xlDown)).Select
ActiveSheet.Range("$D$5:$D$3486").AutoFilter Field:=1, Criteria1:=(min_ano + y - 1)


Range("M1").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Workbooks.Add
Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

If (min_ano + y - 1) >= 2010 Then

extensao = "0" & (min_ano + y - 1 - 2000)

Else

extensao = "00" & (min_ano + y - 1 - 2000)

End If

Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="C:\Users\Murilo\Dropbox\DOUTORADO\WUR\Performance\FieldData\Weather_Files\" & SIG & "." & extensao & "", _
        FileFormat:=xlTextPrinter, CreateBackup:=False
    
    Application.DisplayAlerts = False
    ActiveWindow.Close

Next

Next

End Sub

Sub limpa_plan()

Sheets("ENTRADA").Select
Columns("S:AA").Select
Selection.ClearContents

End Sub

