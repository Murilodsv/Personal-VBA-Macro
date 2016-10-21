Attribute VB_Name = "Módulo62"
Sub Macro40()
Attribute Macro40.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro40 Macro
'

'
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "titulo"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "titulo"
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
    ActiveChart.ChartArea.Select
    ActiveChart.Axes(xlValue).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "comparado"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "comparado"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 9).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 9).Font
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
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlValue).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "eixo y"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "eixo y"
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
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "eixo x"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "eixo x"
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
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    ActiveChart.ChartArea.Select
    Range("W19").Select
End Sub
Sub Macro41()
Attribute Macro41.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro41 Macro
'

'
    ActiveChart.Axes(xlValue).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "eixo y"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "eixo y"
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
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    ActiveChart.ChartArea.Select
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "eixo y"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "eixo y"
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
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    ActiveChart.ChartArea.Select
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "eixo x"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "eixo x"
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
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    Range("W21").Select
End Sub
Sub Macro42()
Attribute Macro42.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro42 Macro
'

'
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).MaximumScale = 600
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MaximumScale = 700
End Sub
Sub Macro43()
Attribute Macro43.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro43 Macro
'

'
    Range("S22").Select
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = 1
End Sub
Sub Macro44()
Attribute Macro44.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro44 Macro
'

'
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.ChartArea.Copy
    Sheets("SAIDA").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    Range("S2").Select
    ActiveSheet.Pictures.Paste.Select
    Range("A2:S2").Select
    Range("S2").Activate
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    Columns("R:R").Select
    ActiveWindow.LargeScroll ToRight:=1
    Columns("R:AE").Select
    Selection.Delete Shift:=xlToLeft
    Range("Q5").Select
    ActiveSheet.Shapes.Range(Array("Picture 1")).Select
    Columns("R:R").Select
    ActiveSheet.Shapes.Range(Array("Picture 1")).Select
    Selection.ShapeRange.IncrementLeft 129
    Selection.ShapeRange.IncrementTop 4.5
    Columns("S:S").Select
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    Columns("S:AF").Select
    Selection.Delete Shift:=xlToLeft
    Columns("R:AB").Select
    Selection.ClearContents
    Selection.ClearContents
    Range("AE9").Select
    ActiveSheet.Shapes.Range(Array("Picture 1")).Select
    Range("AE12").Select
End Sub
Sub Macro45()
Attribute Macro45.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro45 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Picture 1")).Select
    Selection.Delete
End Sub
