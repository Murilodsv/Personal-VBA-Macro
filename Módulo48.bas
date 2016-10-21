Attribute VB_Name = "Módulo48"
Sub Macro29()
Attribute Macro29.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro29 Macro
'

'
    ActiveSheet.Range("$A$6:$A$12059").AutoFilter Field:=1, Criteria1:="<>", _
        Operator:=xlAnd
    ActiveSheet.Range("$A$6:$A$12059").AutoFilter Field:=1, Criteria1:="<>"
End Sub
