Attribute VB_Name = "Módulo4"
Sub Cola_Grafico()
'
'

NLINHA = 3
ano = 1
N_ANO = Sheets("Lista").Range("D" & 2).Value

For x = 1 To N_ANO

Windows("Cola_Grafico.xlsx").Activate
dire = Sheets("Lista").Range("A" & 2).Value
arq = Sheets("Lista").Range("B" & x + 1).Value

Workbooks.Open Filename:="" & dire & "\" & arq & ".xls"
    
    Windows("" & arq & ".xls").Activate
    Sheets("Graf-1-temp_TRIM-JFM").Select
    ActiveSheet.Shapes.Range(Array("Chart 3", "Chart 4")).Select
    Selection.Copy
    Windows("Cola_Grafico.xlsx").Activate
    Sheets("Graficos").Select
    Range(Cells(NLINHA, 3), Cells(NLINHA, 3)).Select
    ActiveSheet.Pictures.Paste.Select
    
    NLINHA = NLINHA + 23
    
    Windows("" & arq & ".xls").Activate
    Sheets("Graf-1-temp_TRIM-AMJ").Select
    ActiveSheet.Shapes.Range(Array("Chart 3", "Chart 4")).Select
    Selection.Copy
    Windows("Cola_Grafico.xlsx").Activate
    Sheets("Graficos").Select
    Range(Cells(NLINHA, 3), Cells(NLINHA, 3)).Select
    ActiveSheet.Pictures.Paste.Select
    
    NLINHA = NLINHA + 23
    
    Windows("" & arq & ".xls").Activate
    Sheets("Graf-1-temp_TRIM-JAS").Select
    ActiveSheet.Shapes.Range(Array("Chart 3", "Chart 4")).Select
    Selection.Copy
    Windows("Cola_Grafico.xlsx").Activate
    Sheets("Graficos").Select
    Range(Cells(NLINHA, 3), Cells(NLINHA, 3)).Select
    ActiveSheet.Pictures.Paste.Select
    
    NLINHA = NLINHA + 23
    
    Windows("" & arq & ".xls").Activate
    Sheets("Graf-1-temp_TRIM-OND").Select
    ActiveSheet.Shapes.Range(Array("Chart 3", "Chart 4")).Select
    Selection.Copy
    Windows("Cola_Grafico.xlsx").Activate
    Sheets("Graficos").Select
    Range(Cells(NLINHA, 3), Cells(NLINHA, 3)).Select
    ActiveSheet.Pictures.Paste.Select
    
    NLINHA = NLINHA + 23
    
    Windows("" & arq & ".xls").Activate
    Sheets("Graf-1-temp_ANO").Select
    ActiveSheet.Shapes.Range(Array("Chart 3", "Chart 4")).Select
    Selection.Copy
    Windows("Cola_Grafico.xlsx").Activate
    Sheets("Graficos").Select
    Range(Cells(NLINHA, 3), Cells(NLINHA, 3)).Select
    ActiveSheet.Pictures.Paste.Select
    
    Range(Cells(ano, 2), Cells(ano, 2)).Select
    ActiveCell.FormulaR1C1 = "" & x + Sheets("Lista").Range("C" & 2).Value - 1 & ""
    
    NLINHA = NLINHA + 25
    ano = ano + 117
    
    Windows("" & arq & ".xls").Activate

    Application.DisplayAlerts = False
    ActiveWindow.Close

Next

End Sub
