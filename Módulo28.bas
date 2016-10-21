Attribute VB_Name = "Módulo28"


Sub OUTPUTS_DSSAT_IMPORTA()
'
' IMPORTA OS ARQUIVOS DE SAIDA OPG, OSW E OEB POR EXPERIMENTO.
'

'

Workbooks.Open Filename:="C:\Murilo\MESTRADO\Simulacao\RESUMO_EXPERIMENTOS.xlsx"

For x = 91 To 120

Windows("RESUMO_EXPERIMENTOS.xlsx").Activate

expe = Sheets("Sequeiro").Range("A" & x + 1).Value
esta = Sheets("Sequeiro").Range("F" & x + 1).Value

Workbooks.Open Filename:="C:\Murilo\MACRO\OUTPUTS_DSSAT_IMPORTA.xlsm"
Application.Calculation = xlManual

    Sheets("OPG").Select
           
    Columns("G:CG").Select
    Selection.ClearContents
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\Murilo\MESTRADO\Simulacao\Batch_DSSAT\Sequence\" & expe & ".OPG", Destination:=Range("$G$1"))
        .Name = "ESAL0001"
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
    
    Sheets("OSW").Select
    Columns("G:BR").Select
    Selection.ClearContents
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\Murilo\MESTRADO\Simulacao\Batch_DSSAT\Sequence\" & expe & ".OSW", Destination:=Range("$G$1"))
        .Name = "ESAL0001"
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
        
    Sheets("OEB").Select
    Columns("F:BR").Select
    Selection.ClearContents
        
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\Murilo\MESTRADO\Simulacao\Batch_DSSAT\Sequence\" & expe & ".OEB", Destination:=Range("$F$1"))
        .Name = "ESAL0001"
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
    
    Columns("A:F").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("OPG").Select
    Columns("A:F").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets("OSW").Select
    Columns("A:F").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("ETP").Select
    Columns("F:F").Select
    Selection.AutoFilter
    ActiveSheet.Range("$F$1:$F$35641").AutoFilter Field:=1, Criteria1:="<>" & esta & ""
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    
    Sheets("OPG").Select
    Range("A14").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$14:$A$20000").AutoFilter Field:=1, Criteria1:="="
    
    Range("A14").Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Selection.AutoFilter
    
    Sheets("OSW").Select
    Range("A13").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$13:$A$20000").AutoFilter Field:=1, Criteria1:="="
    
    Range("A13").Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Selection.AutoFilter
    
    Sheets("OEB").Select
    Range("A11").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$11:$A$20000").AutoFilter Field:=1, Criteria1:="="
    
    Range("A11").Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Selection.AutoFilter
    
    Sheets("MEDIA_TOTAL").Select
    Range(Cells(10, 1), Cells(10, 18)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range(Cells(10, 1), Cells(10, 1)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
           
ActiveWorkbook.SaveAs Filename:="C:\Murilo\MESTRADO\Simulacao\Batch_DSSAT\OUTPUTS_DSSAT\Sequence\" & expe & ".xlsx", _
FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

ActiveWindow.Close


Next

Call OUTPUTS_DSSAT_IMPORTA_TOTAIS

End Sub

Sub OUTPUTS_DSSAT_IMPORTA_TOTAIS()
'
' IMPORTA OS ARQUIVOS DE SAIDA OPG, OSW E OEB POR EXPERIMENTO.
'

'
x = 1
y = 1
K = 91
Workbooks.Open Filename:="C:\Murilo\MESTRADO\Simulacao\RESUMO_EXPERIMENTOS.xlsx"

For y = 1 To 1 ' n solos

linha = 5
Workbooks.Open Filename:="C:\Murilo\MACRO\OUTPUTS_DSSAT_IMPORTA_TOTAIS.xlsm"

For x = 1 To 30 ' n estacoes

Windows("RESUMO_EXPERIMENTOS.xlsx").Activate

expe = Sheets("Sequeiro").Range("A" & K + 1).Value
esta = Sheets("Sequeiro").Range("F" & K + 1).Value
solo = Sheets("Sequeiro").Range("B" & K + 1).Value


Workbooks.Open Filename:="C:\Murilo\MESTRADO\Simulacao\Batch_DSSAT\OUTPUTS_DSSAT\Sequence\" & expe & ".xlsx"


    Sheets("TOTAL").Select
    Range(Cells(5, 1), Cells(36, 72)).Select
    Selection.Copy
   
    Windows("OUTPUTS_DSSAT_IMPORTA_TOTAIS.xlsm").Activate
    Sheets("TOTAL").Select
    Range(Cells(linha, 1), Cells(linha, 72)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Windows("" & expe & ".xlsx").Activate
    Sheets("MEDIA_TOTAL").Select
    Range(Cells(5, 1), Cells(5, 58)).Select
    Selection.Copy
    
    Windows("OUTPUTS_DSSAT_IMPORTA_TOTAIS.xlsm").Activate
    Sheets("MEDIA_TOTAL").Select
    Range(Cells(x + 4, 1), Cells(x + 4, 58)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Windows("" & expe & ".xlsx").Activate
    Sheets("MEDIA_TOTAL").Select
    Range(Cells(9, 1), Cells(9, 2)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("OUTPUTS_DSSAT_IMPORTA_TOTAIS.xlsm").Activate
    Sheets("MEDIA_CICLO").Select
    Range(Cells(3, 1), Cells(3, 1)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Windows("" & expe & ".xlsx").Activate
    Sheets("MEDIA_TOTAL").Select
    Range(Cells(9, 3), Cells(9, 3)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("OUTPUTS_DSSAT_IMPORTA_TOTAIS.xlsm").Activate
    Sheets("MEDIA_CICLO").Select
    Range(Cells(3, x + 2), Cells(3, x + 2)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Windows("" & expe & ".xlsx").Activate
    Sheets("MEDIA_TOTAL").Select
    Range("U10:AC10").Select
    Selection.Copy
    
    Windows("OUTPUTS_DSSAT_IMPORTA_TOTAIS.xlsm").Activate
    Sheets("MEDIA_TOTAL").Select
    Range(Cells(x + 4, 60), Cells(x + 4, 60)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Windows("" & expe & ".xlsx").Activate
    Application.DisplayAlerts = False
    ActiveWindow.Close
    
    K = K + 1
    linha = linha + 32
    
Next


ActiveWorkbook.SaveAs Filename:="C:\Murilo\MESTRADO\Simulacao\Batch_DSSAT\OUTPUTS_DSSAT\Sequence\TOTAIS_" & expe & "_" & solo & ".xlsx", _
FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
ActiveWindow.Close

Next

Call PlayWAV

End Sub

Sub PlayWAV()


    WAVFile = "Som1.wav"
    dire = "C:\Murilo\MACRO"
    WAVFile = dire & "\" & WAVFile
    'For x = 1 To 3
    Call PlaySound(WAVFile, 0&, SND_ASYNC Or SND_FILENAME)
    'Call Sleep(24000) 'timer em milisegundos
    'Next
    'Call PlaySound(False, 0&, SND_ASYNC Or SND_FILENAME)'desliga o som
    
End Sub




