Attribute VB_Name = "Módulo33"
Sub Convert_W_SWAP()
Attribute Convert_W_SWAP.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Format weather data to SWAP weather files format
'

'

For x = 1 To 10

n_anos = Sheets("LISTA").Range("A" & x).Value


For y = 1 To n_anos

Next




Next


    Range("C3").Select
    ActiveCell.FormulaR1C1 = "abc"
    Range("L5").Select
    Cells.find(What:="abc", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
End Sub
