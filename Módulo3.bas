Attribute VB_Name = "Módulo3"
Sub Solver()
Attribute Solver.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SOLVER Macro
'

'
For x = 1 To 12

mes = Sheets("DIARIO").Range("W" & x).Value

Sheets("" & mes & "").Select

    SolverOk SetCell:="$AA$2", MaxMinVal:=2, ValueOf:=0, ByChange:="$X$2:$Z$2", _
        Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverOk SetCell:="$AA$2", MaxMinVal:=2, ValueOf:=0, ByChange:="$X$2:$Z$2", _
        Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverSolve True
    
    Next
    
End Sub
