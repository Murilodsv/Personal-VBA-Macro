Attribute VB_Name = "Módulo49"
Sub limpa()

For x = 1 To 11


Windows("controle.xlsx").Activate
Sheets("lista").Select
sigla = Sheets("lista").Range("M" & x + 1).Value

Windows("Stage1ModelOutputTemplate.xlsx").Activate
Sheets("" & sigla & "").Select

Range("E3:R10000").Select
Selection.ClearContents


Next

End Sub

Sub Macro30()
Attribute Macro30.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro30 Macro
'

'C:\DSSAT45\Seasonal

For x = 1 To 11

Windows("controle.xlsx").Activate
Sheets("lista").Select
wth = Sheets("lista").Range("N" & x + 1).Value
expe = Sheets("lista").Range("M" & x + 1).Value

Sheets("importa").Select
Range("E1").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\DSSAT45\Sugarcane\" & wth & ".opg", Destination:=Range("$E$1"))
        .Name = "ESSP1401_1"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
Calculate
ActiveSheet.Range("$B$14:$B$40392").AutoFilter Field:=1, Criteria1:="<>"

Range("S15").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("Stage1ModelOutputTemplate.xlsx").Activate
Sheets("" & expe & "").Select
Range("F3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'1-----------------

Windows("controle.xlsx").Activate
Sheets("importa").Select

Range("M15").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("Stage1ModelOutputTemplate.xlsx").Activate
Sheets("" & expe & "").Select
Range("G3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'2-----------------

Windows("controle.xlsx").Activate
Sheets("importa").Select

Range("AO15").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("Stage1ModelOutputTemplate.xlsx").Activate
Sheets("" & expe & "").Select
Range("H3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'3----------------

Windows("controle.xlsx").Activate
Sheets("importa").Select

Range("X15").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("Stage1ModelOutputTemplate.xlsx").Activate
Sheets("" & expe & "").Select
Range("I3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
'4------------------

Windows("controle.xlsx").Activate
Sheets("importa").Select

Range("T15").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("Stage1ModelOutputTemplate.xlsx").Activate
Sheets("" & expe & "").Select
Range("J3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'5-----------------

Windows("controle.xlsx").Activate
Sheets("importa").Select

Range("W15").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("Stage1ModelOutputTemplate.xlsx").Activate
Sheets("" & expe & "").Select
Range("K3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'6------------------


Windows("controle.xlsx").Activate
Sheets("importa").Select

Range("AA15").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("Stage1ModelOutputTemplate.xlsx").Activate
Sheets("" & expe & "").Select
Range("L3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'7------------------

Windows("controle.xlsx").Activate
Sheets("importa").Select

Range("N15").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("Stage1ModelOutputTemplate.xlsx").Activate
Sheets("" & expe & "").Select
Range("M3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'8-----------------

Windows("controle.xlsx").Activate
Sheets("importa").Select

Range("R15").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("Stage1ModelOutputTemplate.xlsx").Activate
Sheets("" & expe & "").Select
Range("N3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'9-------------------

Windows("controle.xlsx").Activate
Sheets("importa").Select

Range("Q15").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("Stage1ModelOutputTemplate.xlsx").Activate
Sheets("" & expe & "").Select
Range("O3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'10--------------------

Windows("controle.xlsx").Activate
Sheets("importa").Select

Range("O15").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("Stage1ModelOutputTemplate.xlsx").Activate
Sheets("" & expe & "").Select
Range("P3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'11---------------------
        
Windows("controle.xlsx").Activate
Sheets("importa").Select

Range("P15").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("Stage1ModelOutputTemplate.xlsx").Activate
Sheets("" & expe & "").Select
Range("Q3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'12-------------------------

Windows("controle.xlsx").Activate
Sheets("importa").Select

Range("M15").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("Stage1ModelOutputTemplate.xlsx").Activate
Sheets("" & expe & "").Select
Range("R3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'13------------------------treatment

Windows("controle.xlsx").Activate
Sheets("importa").Select

Range("B15").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("Stage1ModelOutputTemplate.xlsx").Activate
Sheets("" & expe & "").Select
Range("C3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'14------------------------DAP

Windows("controle.xlsx").Activate
Sheets("importa").Select

Range("I15").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("Stage1ModelOutputTemplate.xlsx").Activate
Sheets("" & expe & "").Select
Range("E3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("controle.xlsx").Activate
Sheets("importa").Select

Application.DisplayAlerts = False
Columns("E:BR").Select
Selection.ClearContents

Calculate
  
Next

End Sub

Sub APSIM_OUT_AGMIP()
'
' Macro30 Macro
'

'C:\DSSAT45\Seasonal

linha = 3
NLINHA = 1
For x = 1 To 2

Windows("controle_APSIM.xlsx").Activate
Sheets("lista").Select
wth = Sheets("APSIM").Range("A" & x + 1).Value
expe = Sheets("APSIM").Range("B" & x + 1).Value
TRAT = Sheets("APSIM").Range("C" & x + 1).Value


Windows("controle_APSIM.xlsx").Activate
Sheets("importa").Select

ActiveSheet.Range("$B$4:$B$4153").AutoFilter Field:=1

Application.DisplayAlerts = False
Columns("E:PT").Select
Selection.ClearContents

Sheets("importa").Select
Range("E1").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\Murilo\DOUTORADO\AGMIP\DataSubmission_Stage1_v2\APSIM\Simulacoes_4\Simulacoes_4\" & wth & "", Destination:=Range("$E$1"))
        .Name = "ESSP1401_1"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
Calculate

If TRAT = 1 Then

linha = 3

Else

linha = linha + NLINHA

End If

Sheets("APSIM").Select
Range(Cells(x + 1, 4), Cells(x + 1, 6)).Select
Selection.Copy

Sheets("importa").Select
Range("B2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Calculate

NLINHA = Sheets("APSIM").Range("H" & x + 1).Value

ActiveSheet.Range("$B$4:$B$4153").AutoFilter Field:=1, Criteria1:="1"

Range("D4:PT4").Select
Selection.ClearContents

For y = 1 To 14

Windows("controle_APSIM.xlsx").Activate
VARI = Sheets("APSIM").Range("I" & y + 1).Value

Rows("3:3").Select
    Selection.find(What:=VARI, After:=ActiveCell, LookIn:=xlFormulas _
        , LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("Stage1ModelOutputTemplate_APSIM.xlsx").Activate
Sheets("" & expe & "").Select
Range(Cells(linha, y + 4), Cells(linha, y + 4)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Next

Windows("controle_APSIM.xlsx").Activate
Range("D4").Select
Selection.End(xlDown).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("Stage1ModelOutputTemplate_APSIM.xlsx").Activate
Sheets("" & expe & "").Select
Range(Cells(linha, 3), Cells(linha, 3)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("controle_APSIM.xlsx").Activate
Sheets("importa").Select

ActiveSheet.Range("$B$4:$B$4153").AutoFilter Field:=1

Application.DisplayAlerts = False
Columns("E:PT").Select
Selection.ClearContents

Calculate
  
Next

End Sub


