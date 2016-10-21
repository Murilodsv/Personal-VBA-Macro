Attribute VB_Name = "Módulo30"
Sub Cria_experimento()
'
' Cria arquivos experimentais DSSAT45
'

'

Workbooks.Open Filename:="C:\Murilo\MESTRADO\Simulacao\RESUMO_EXPERIMENTOS.xlsx"

For x = 1 To 30

Windows("RESUMO_EXPERIMENTOS.xlsx").Activate

cenario = "Potencial"

expe = Sheets("" & cenario & "").Range("A" & x + 1).Value
solo = Sheets("" & cenario & "").Range("B" & x + 1).Value
plan = Sheets("" & cenario & "").Range("C" & x + 1).Value
colh = Sheets("" & cenario & "").Range("D" & x + 1).Value
cult = Sheets("" & cenario & "").Range("E" & x + 1).Value
esta = Sheets("" & cenario & "").Range("F" & x + 1).Value
ssim = Sheets("" & cenario & "").Range("G" & x + 1).Value
inic = Sheets("" & cenario & "").Range("H" & x + 1).Value
inin = Sheets("" & cenario & "").Range("I" & x + 1).Value

Workbooks.Open Filename:="C:\Murilo\MESTRADO\Simulacao\IMPORTA.xlsx"

    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\Murilo\MACRO\MODELO.SQX", Destination:=Range("$A$1"))
        .Name = "ESAL9999"
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
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
   
    Columns("A:A").Select
    
    Selection.Replace What:="$expe$", Replacement:=expe, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Selection.Replace What:="$solo$", Replacement:=solo, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Selection.Replace What:="$plan$", Replacement:=plan, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Selection.Replace What:="$colh$", Replacement:=colh, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Selection.Replace What:="$cult$", Replacement:=cult, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Selection.Replace What:="$esta$", Replacement:=esta, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Selection.Replace What:="$ssim$", Replacement:=ssim, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Selection.Replace What:="$inic$", Replacement:=inic, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Selection.Replace What:="$inin$", Replacement:=inin, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    ActiveWorkbook.SaveAs Filename:="C:\DSSAT45\Sequence\" & expe & ".SQX", _
        FileFormat:=xlTextPrinter, CreateBackup:=False
    
    Application.DisplayAlerts = False
    ActiveWindow.Close
        
               
Next
        
    Call Cria_v45
        
End Sub

Sub Cria_v45()
'
' Macro25 Macro
'

'

Workbooks.Open Filename:="C:\Murilo\MESTRADO\Simulacao\RESUMO_EXPERIMENTOS.xlsx"

For x = 1 To 30

Windows("RESUMO_EXPERIMENTOS.xlsx").Activate

cenario = "Potencial"

expe = Sheets("" & cenario & "").Range("A" & x + 1).Value
solo = Sheets("" & cenario & "").Range("B" & x + 1).Value
plan = Sheets("" & cenario & "").Range("C" & x + 1).Value
colh = Sheets("" & cenario & "").Range("D" & x + 1).Value
cult = Sheets("" & cenario & "").Range("E" & x + 1).Value
esta = Sheets("" & cenario & "").Range("F" & x + 1).Value
ssim = Sheets("" & cenario & "").Range("G" & x + 1).Value
inic = Sheets("" & cenario & "").Range("H" & x + 1).Value
inin = Sheets("" & cenario & "").Range("I" & x + 1).Value


Workbooks.Open Filename:="C:\Murilo\MESTRADO\Simulacao\IMPORTA.xlsx"

    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\Murilo\MACRO\MODELO_Q.v45", Destination:=Range("$A$1"))
        .Name = "ESAL9999"
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
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
   
    Columns("A:A").Select
    
    Selection.Replace What:="$expe$", Replacement:=expe, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    
    ActiveWorkbook.SaveAs Filename:="C:\Murilo\MESTRADO\Simulacao\Batch_DSSAT\Sequence\" & expe & ".v45", _
        FileFormat:=xlTextPrinter, CreateBackup:=False
    
    Application.DisplayAlerts = False
    ActiveWindow.Close
        
           
Next
        
        
End Sub

Sub abre()

'Abre um determinado arquivo pelo cmd.exe
'/c para fechar no final e /k para manter aberto no final
'START comando do cmd para abrir um arquivo
'/MIN minimizado /MAX maximizado

    Shell "C:\Windows\System32\cmd.exe """ & "/c START /Dc:\Murilo\MESTRADO\Simulacao\Batch_DSSAT\Sequence\ /MAX BATCH_DSSAT.bat " & "", vbNormalFocus
           
End Sub


 



