Sub QGIS_Batch()
'
' QGIS_Batch Macro
' Creating QGIS batch file for Zonal Statistics in several shapefiles

'

L = 134
iL = L + 1

    For x = 2 To 28
    
    tif = Sheets("list").Range("A" & x).Value
    
    Sheets("Planilha1").Select
    Range(Cells(1, 1), Cells(L, 1)).Select
    Selection.Copy
    
    Range(Cells(iL, 1), Cells(iL, 1)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range(Cells(iL, 1), Cells(iL + L, 1)).Select
    Selection.Replace What:="BLDFIE_M_sl1_250m", Replacement:= _
        tif, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:= _
        False, SearchFormat:=False, ReplaceFormat:=False
        
    iL = iL + L + 1
    
    Next
    
End Sub
