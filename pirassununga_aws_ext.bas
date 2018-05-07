Sub wth_pirassununga()
'
' Pirassununga WTH creation for Crop Modelling
'-21.961517, -47.470232
        'from : http://www.agrariasusp.com.br/agrariasusp01/estacao.html

'
Dim m(12) As String

m(0) = "jan"
m(1) = "fev"
m(2) = "mar"
m(3) = "abr"
m(4) = "mai"
m(5) = "jun"
m(6) = "jul"
m(7) = "ago"
m(8) = "set"
m(9) = "out"
m(10) = "nov"
m(11) = "dez"

yr = 12
l = 3
For x = 1 To 6

Workbooks.Open Filename:="C:\Users\PC-600\Dropbox (Farmers Edge)\MuriloVianna\Modeling\Ferrari\WTH\Raw\extract_data.xlsx"

rawdata = Sheets("list").Range("A" & x).Value
Workbooks.Open Filename:="C:\Users\PC-600\Dropbox (Farmers Edge)\MuriloVianna\Modeling\Ferrari\WTH\Raw\" & rawdata & ".xls"

For y = 1 To 12

Windows("" & rawdata & ".xls").Activate
Sheets("" & m(y - 1) & "" & yr & "").Select
Range(Cells(3, 1), Cells(4500, 45)).Select
Selection.Copy

Windows("extract_data.xlsx").Activate
Sheets("hourly").Select
Range(Cells(l, 1), Cells(l, 1)).Select

Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

l = l + 4500

Next

yr = yr + 1

Windows("" & rawdata & ".xls").Activate
Application.DisplayAlerts = False
ActiveWindow.Close

Next
    
    

End Sub
