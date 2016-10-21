Attribute VB_Name = "Módulo8"
Sub Exercicio5_AeDP()
Attribute Exercicio5_AeDP.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro5 Macro
'

'

Windows("Ex#05_DPM.xlsx").Activate
    
For x = 1 To 55
    
Sheets("Dados_2005").Select
Range(Cells(x + 3, 29), Cells(x + 3, 29 + 23)).Select ' Tar
Selection.Copy

Sheets("Planilha de Cálculo da DPM").Select
Range("C4").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True

Sheets("Dados_2005").Select
Range(Cells(x + 61, 29), Cells(x + 61, 29 + 23)).Select ' Rn
Selection.Copy

Sheets("Planilha de Cálculo da DPM").Select
Range("B4").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True

Sheets("Dados_2005").Select
Range(Cells(x + 119, 29), Cells(x + 119, 29 + 23)).Select ' UR
Selection.Copy

Sheets("Planilha de Cálculo da DPM").Select
Range("D4").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True

Sheets("Dados_2005").Select
Range(Cells(x + 177, 29), Cells(x + 177, 29 + 23)).Select ' PPT
Selection.Copy

Sheets("Planilha de Cálculo da DPM").Select
Range("F4").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True

Sheets("Dados_2005").Select
Range(Cells(x + 236, 29), Cells(x + 236, 29 + 23)).Select ' U
Selection.Copy

Sheets("Planilha de Cálculo da DPM").Select
Range("E4").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True

Range("Y29,AA29,AE29,AK29").Select
Selection.Copy
Sheets("OBS").Select
Range(Cells(x + 2, 7), Cells(x + 2, 7)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Sheets("Planilha de Cálculo da DPM").Select
Range("B4").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

Next
    
    
    
End Sub

Sub Macro6()
Attribute Macro6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro6 Macro
'

'
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "TESTE"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "TESTE"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 5).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 5).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    Range("Z18").Select
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "TESTE1"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "TESTE1"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 6).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 6).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    Range("AD7").Select
End Sub
Sub Macro7()
Attribute Macro7.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro7 Macro
'

'
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.SeriesCollection(1).Trendlines(3).Select
    Selection.Delete
    ActiveSheet.ChartObjects("Gráfico 1""").Activate
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Obsersdfsddo"
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "wdfwservado (kg/ha)"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Obsersdfsddo"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 12).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 12).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    Range("Y18").Select
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.Axes(xlValue).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "wdfwservado (kg/ha)"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "wdfwservado (kg/ha)"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 19).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 19).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    Range("W23").Select
End Sub
Sub Macro8()
Attribute Macro8.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro8 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Picture 9")).Select
    Selection.Delete
End Sub
Sub Macro9()
Attribute Macro9.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro9 Macro
'

'
    ActiveChart.ChartArea.Select
    Application.CutCopyMode = False
    ActiveChart.ChartArea.Copy
End Sub
Sub Macro10()
Attribute Macro10.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro10 Macro
'

'
    Range("AC22").Select
    ActiveSheet.Shapes.Range(Array("Picture 10")).Select
    Selection.Delete
End Sub
Sub Macro11()
Attribute Macro11.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro11 Macro
'

'
    Range("X22").Select
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.ChartArea.Copy
    Sheets("SAIDA").Select
    Range("W21").Select
    Sheets("BASE_ESTAT").Select
    Range("V22").Select
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.ChartArea.Copy
    Sheets("SAIDA").Select
    ActiveWindow.SmallScroll Down:=12
    Range("W37").Select
    ActiveSheet.Paste
End Sub
Sub Macro12()
Attribute Macro12.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro12 Macro
'

'
    
    Sheets("Plan8").Name = "WEF"
End Sub
Sub Macro13()
Attribute Macro13.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro13 Macro
'

'
    Columns("V:AD").Select
    Selection.Cut
    Sheets("Plan12").Select
    Columns("A:A").Select
    ActiveSheet.Paste
    Sheets("Plan12").Select
    ActiveWindow.SelectedSheets.Delete
End Sub
Sub Macro14()
Attribute Macro14.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro14 Macro
'

'

x = 60

    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.SeriesCollection(1).XValues = "=BASE_ESTAT!$A$6:$A$" & x & ""
    ActiveChart.SeriesCollection(1).Values = "=BASE_ESTAT!$B$6:$B$" & x & ""
   
End Sub
