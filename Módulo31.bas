Attribute VB_Name = "Módulo31"
Sub Solver_Theta()
'
' SOLVER Macro
'

l = 314

    
For x = 1 To 5

    SolverOk SetCell:=Range("R" & l), MaxMinVal:=3, ValueOf:=0, ByChange:=Range(Cells(l, 14), Cells(l, 16)), _
        Engine:=1, EngineDesc:="GRG Nonlinear"
        '"GRG Nonlinear"
        
    'SolverChange CellRef:=Range("D" & X), Relation:=1, _
        FormulaText:=Range("E" & X)
    'SolverChange CellRef:=Range("D" & X), Relation:=3, _
        FormulaText:=0
    
    SolverSolve True
    
    l = l + (346 - 314)
    
    Next
    
End Sub


