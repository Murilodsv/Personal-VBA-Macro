Attribute VB_Name = "Módulo19"
Sub EXTRAI_ACUMULADO_OPG()
'
' Extrai valores acumulados de arquivos .opg de saída do DSSAT.
' Macro desenvolvida em 16/06/2011 por Murilo Santos Vianna
'

'

Windows("EXTRAI_ACUMULADO_OPG.xlsm").Activate
N_OPG = Sheets("PARAMETROS").Range("F" & 2).Value
N_SAIDA = Sheets("PARAMETROS").Range("G" & 2).Value
K = 1
For x = 1 To N_OPG

ARQUIVO_ENTRADA = Sheets("PARAMETROS").Range("B" & x + 1).Value

If N_SAIDA = 1 Then
NOME_SAIDA = Sheets("PARAMETROS").Range("D" & 2).Value
DIRETORIO_SAIDA = Sheets("PARAMETROS").Range("C" & 2).Value
Else
NOME_SAIDA = Sheets("PARAMETROS").Range("D" & x + 1).Value
DIRETORIO_SAIDA = Sheets("PARAMETROS").Range("C" & x + 1).Value
End If

DIRETORIO_ENTRADA = Sheets("PARAMETROS").Range("A" & x + 1).Value
DIRETORIO_PLANILHA = Sheets("PARAMETROS").Range("E" & 2).Value

Sheets("IMPORTA").Select

 With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & DIRETORIO_ENTRADA & "\" & ARQUIVO_ENTRADA & "", Destination:=Range("A1"))
        .Name = "1_39"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 932
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With

Columns("E:F").Select
Selection.Replace What:="=-", Replacement:="", LookAt:=xlPart, _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False

N_RUNS = Sheets("PARAMETROS").Range("H" & 2).Value

linha = 1
L_ESTACAO = 5
L_MODELO = 6
L_EXPERIMENTO = 7
L_TRATAMENTO = 9
L_RUN = 5
L_SIMULACAO = 1

For y = 1 To N_RUNS

Sheets("IMPORTA").Select
Range(Cells(linha + 1, 5), Cells(linha + 1, 5)).Select
Selection.End(xlDown).Select
Selection.End(xlDown).Select
Selection.End(xlDown).Select
Selection.End(xlDown).Select
Selection.End(xlDown).Select
Selection.Copy

Sheets("PARAMETROS").Select
Range("I2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False

L_SIMULACAO = Sheets("PARAMETROS").Range("J" & 2).Value + 13
linha = linha + L_SIMULACAO

Sheets("IMPORTA").Select
Range(Cells(linha, 2), Cells(linha, 2)).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Copy

Sheets("RESULTADO").Select
Range(Cells(K + 4, 7), Cells(K + 4, 7)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False

ESPACO_100 = Sheets("IMPORTA").Range("C" & L_ESTACAO + 3).Value

If ESPACO_100 = ":" Then

ESTACAO = Sheets("IMPORTA").Range("D" & L_ESTACAO).Value
modelo = Sheets("IMPORTA").Range("D" & L_MODELO).Value
experimento = Sheets("IMPORTA").Range("D" & L_EXPERIMENTO).Value
TRATAMENTO = Sheets("IMPORTA").Range("E" & L_TRATAMENTO).Value
Runs = Sheets("IMPORTA").Range("B" & L_RUN).Value

Else

ESTACAO = Sheets("IMPORTA").Range("E" & L_ESTACAO).Value
modelo = Sheets("IMPORTA").Range("D" & L_MODELO).Value
experimento = Sheets("IMPORTA").Range("D" & L_EXPERIMENTO).Value
TRATAMENTO = Sheets("IMPORTA").Range("E" & L_TRATAMENTO).Value
Runs = Sheets("IMPORTA").Range("B" & L_RUN).Value

End If

Sheets("RESULTADO").Select
Range(Cells(K + 4, 1), Cells(K + 4, 1)).Select
ActiveCell.FormulaR1C1 = ESTACAO
Range(Cells(K + 4, 2), Cells(K + 4, 2)).Select
ActiveCell.FormulaR1C1 = modelo
Range(Cells(K + 4, 3), Cells(K + 4, 3)).Select
ActiveCell.FormulaR1C1 = experimento
Range(Cells(K + 4, 4), Cells(K + 4, 4)).Select
ActiveCell.FormulaR1C1 = TRATAMENTO
Range(Cells(K + 4, 5), Cells(K + 4, 5)).Select
ActiveCell.FormulaR1C1 = Runs
Range(Cells(K + 4, 6), Cells(K + 4, 6)).Select
ActiveCell.FormulaR1C1 = ARQUIVO_ENTRADA


L_ESTACAO = L_ESTACAO + L_SIMULACAO
L_MODELO = L_MODELO + L_SIMULACAO
L_EXPERIMENTO = L_EXPERIMENTO + L_SIMULACAO
L_TRATAMENTO = L_TRATAMENTO + L_SIMULACAO
L_RUN = L_RUN + L_SIMULACAO
K = K + 1

Next

Sheets("IMPORTA").Select
Columns("A:IV").Select
Selection.ClearContents

'If N_SAIDA > 1 Then

'Sheets("RESULTADO").Select
'ActiveWorkbook.SaveAs Filename:= _
        "" & DIRETORIO_SAIDA & "\" & NOME_SAIDA & "", FileFormat:=xlNormal, _
        Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, _
        CreateBackup:=False

'Workbooks.Open Filename:="" & DIRETORIO_PLANILHA & "\EXTRAI_ACUMULADO_OPG.XLS"

'Windows("" & NOME_SAIDA & ".XLS").Activate
'ActiveWindow.Close

'Else

'End If


Next

Sheets("RESULTADO").Select
ActiveWorkbook.SaveAs Filename:= _
        "" & DIRETORIO_SAIDA & "\" & NOME_SAIDA & "", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
Application.DisplayAlerts = False
Sheets("RESULTADO").Select

End Sub






