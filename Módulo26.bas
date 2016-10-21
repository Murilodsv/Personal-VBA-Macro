Attribute VB_Name = "Módulo26"
Sub graficos()
Attribute graficos.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro23 Macro
'

'

'---------------OPG---------------
Sheets("OPG").Select
GRAF = Sheets("OPG").Range("AY" & 6).Value
X1 = Sheets("OPG").Range("AZ" & 6).Value
X2 = Sheets("OPG").Range("BA" & 6).Value
Y1 = Sheets("OPG").Range("BB" & 6).Value
Y2 = Sheets("OPG").Range("BC" & 6).Value

ActiveSheet.ChartObjects("" & GRAF & "").Activate

ActiveChart.SeriesCollection(1).XValues = Sheets("OPG").Range(Cells(Y1, X1 - 3), Cells(Y2, X2 - 3))
ActiveChart.SeriesCollection(2).XValues = Sheets("OPG").Range(Cells(Y1, X1 - 3), Cells(Y2, X2 - 3))
ActiveChart.SeriesCollection(3).XValues = Sheets("OPG").Range(Cells(Y1, X1 - 3), Cells(Y2, X2 - 3))

    
ActiveChart.SeriesCollection(1).Values = Sheets("OPG").Range(Cells(Y1, X1 + 5), Cells(Y2, X2 + 5))
ActiveChart.SeriesCollection(3).Values = Sheets("OPG").Range(Cells(Y1, X1 + 6), Cells(Y2, X2 + 6))
ActiveChart.SeriesCollection(2).Values = Sheets("OPG").Range(Cells(Y1, X1 + 41), Cells(Y2, X2 + 41))


ActiveChart.SeriesCollection(1).Name = "=""TCH (t.ha-¹)"""
ActiveChart.SeriesCollection(3).Name = "=""TCH SECO(t.ha-¹)"""
ActiveChart.SeriesCollection(2).Name = "=""POL (%)"""


Sheets("OPG").Select
GRAF = Sheets("OPG").Range("AY" & 7).Value
X1 = Sheets("OPG").Range("AZ" & 7).Value
X2 = Sheets("OPG").Range("BA" & 7).Value
Y1 = Sheets("OPG").Range("BB" & 7).Value
Y2 = Sheets("OPG").Range("BC" & 7).Value

ActiveSheet.ChartObjects("" & GRAF & "").Activate

ActiveChart.SeriesCollection(1).XValues = Sheets("OPG").Range(Cells(Y1, X1 - 3), Cells(Y2, X2 - 3))
    
ActiveChart.SeriesCollection(1).Values = Sheets("OPG").Range(Cells(Y1, X1 + 15), Cells(Y2, X2 + 15))

ActiveChart.SeriesCollection(1).Name = "=""WSPD"""

Sheets("OPG").Select
GRAF = Sheets("OPG").Range("AY" & 8).Value
X1 = Sheets("OPG").Range("AZ" & 8).Value
X2 = Sheets("OPG").Range("BA" & 8).Value
Y1 = Sheets("OPG").Range("BB" & 8).Value
Y2 = Sheets("OPG").Range("BC" & 8).Value

ActiveSheet.ChartObjects("" & GRAF & "").Activate

ActiveChart.SeriesCollection(1).XValues = Sheets("OPG").Range(Cells(Y1, X1 - 3), Cells(Y2, X2 - 3))
    
ActiveChart.SeriesCollection(1).Values = Sheets("OPG").Range(Cells(Y1, X1 + 16), Cells(Y2, X2 + 16))

ActiveChart.SeriesCollection(1).Name = "=""WSGD"""

Sheets("OPG").Select
GRAF = Sheets("OPG").Range("AY" & 9).Value
X1 = Sheets("OPG").Range("AZ" & 9).Value
X2 = Sheets("OPG").Range("BA" & 9).Value
Y1 = Sheets("OPG").Range("BB" & 9).Value
Y2 = Sheets("OPG").Range("BC" & 9).Value

ActiveSheet.ChartObjects("" & GRAF & "").Activate

ActiveChart.SeriesCollection(1).XValues = Sheets("OPG").Range(Cells(Y1, X1 - 3), Cells(Y2, X2 - 3))
    
ActiveChart.SeriesCollection(1).Values = Sheets("OPG").Range(Cells(Y1, X1 + 17), Cells(Y2, X2 + 17))

ActiveChart.SeriesCollection(1).Name = "=""SW30"""



For OPG = 5 To 10

GRAF = Sheets("OPG").Range("AY" & OPG + 5).Value
X1 = Sheets("OPG").Range("AZ" & OPG + 5).Value
X2 = Sheets("OPG").Range("BA" & OPG + 5).Value
Y1 = Sheets("OPG").Range("BB" & OPG + 5).Value
Y2 = Sheets("OPG").Range("BC" & OPG + 5).Value

   ActiveSheet.ChartObjects("" & GRAF & "").Activate
    
    ActiveChart.SeriesCollection(1).XValues = Sheets("OPG").Range(Cells(Y1, X1), Cells(Y2, X2))
    ActiveChart.SeriesCollection(2).XValues = Sheets("OPG").Range(Cells(Y1, X1), Cells(Y2, X2))
    ActiveChart.SeriesCollection(3).XValues = Sheets("OPG").Range(Cells(Y1, X1), Cells(Y2, X2))
    ActiveChart.SeriesCollection(4).XValues = Sheets("OPG").Range(Cells(Y1, X1), Cells(Y2, X2))
    
    ActiveChart.SeriesCollection(4).Values = Sheets("OPG").Range(Cells(Y1 + 3, X1), Cells(Y2 + 3, X2))
    ActiveChart.SeriesCollection(3).Values = Sheets("OPG").Range(Cells(Y1 + 2, X1), Cells(Y2 + 2, X2))
    ActiveChart.SeriesCollection(2).Values = Sheets("OPG").Range(Cells(Y1 + 1, X1), Cells(Y2 + 1, X2))
    ActiveChart.SeriesCollection(1).Values = Sheets("OPG").Range(Cells(Y1 + 5, X1), Cells(Y2 + 5, X2))
    
    
   Next


'---------------OEB---------------
Sheets("OEB").Select
GRAF = Sheets("OEB").Range("AD" & 6).Value
X1 = Sheets("OEB").Range("AE" & 6).Value
X2 = Sheets("OEB").Range("AF" & 6).Value
Y1 = Sheets("OEB").Range("AG" & 6).Value
Y2 = Sheets("OEB").Range("AH" & 6).Value

ActiveSheet.ChartObjects("" & GRAF & "").Activate

ActiveChart.SeriesCollection(1).XValues = Sheets("OEB").Range(Cells(Y1, X1), Cells(Y2, X2))
ActiveChart.SeriesCollection(2).XValues = Sheets("OEB").Range(Cells(Y1, X1), Cells(Y2, X2))

    
ActiveChart.SeriesCollection(1).Values = Sheets("OEB").Range(Cells(Y1, X1 + 14), Cells(Y2, X2 + 14))
ActiveChart.SeriesCollection(2).Values = Sheets("OEB").Range(Cells(Y1, X1 + 15), Cells(Y2, X2 + 15))


ActiveChart.SeriesCollection(1).Name = "=""EOAC"""
ActiveChart.SeriesCollection(2).Name = "=""ETAC"""


Sheets("OEB").Select
GRAF = Sheets("OEB").Range("AD" & 7).Value
X1 = Sheets("OEB").Range("AE" & 7).Value
X2 = Sheets("OEB").Range("AF" & 7).Value
Y1 = Sheets("OEB").Range("AG" & 7).Value
Y2 = Sheets("OEB").Range("AH" & 7).Value

ActiveSheet.ChartObjects("" & GRAF & "").Activate

ActiveChart.SeriesCollection(1).XValues = Sheets("OEB").Range(Cells(Y1, X1), Cells(Y2, X2))
ActiveChart.SeriesCollection(2).XValues = Sheets("OEB").Range(Cells(Y1, X1), Cells(Y2, X2))
    
ActiveChart.SeriesCollection(1).Values = Sheets("OEB").Range(Cells(Y1, X1 + 6), Cells(Y2, X2 + 6))
ActiveChart.SeriesCollection(2).Values = Sheets("OEB").Range(Cells(Y1, X1 + 9), Cells(Y2, X2 + 9))


ActiveChart.SeriesCollection(1).Name = "=""EOAA"""
ActiveChart.SeriesCollection(2).Name = "=""ETAA"""

For OEB = 3 To 9

GRAF = Sheets("OEB").Range("AD" & OEB + 5).Value
X1 = Sheets("OEB").Range("AE" & OEB + 5).Value
X2 = Sheets("OEB").Range("AF" & OEB + 5).Value
Y1 = Sheets("OEB").Range("AG" & OEB + 5).Value
Y2 = Sheets("OEB").Range("AH" & OEB + 5).Value

   ActiveSheet.ChartObjects("" & GRAF & "").Activate
    
    ActiveChart.SeriesCollection(1).XValues = Sheets("OEB").Range(Cells(Y1, X1), Cells(Y2, X2))
    ActiveChart.SeriesCollection(2).XValues = Sheets("OEB").Range(Cells(Y1, X1), Cells(Y2, X2))
    ActiveChart.SeriesCollection(3).XValues = Sheets("OEB").Range(Cells(Y1, X1), Cells(Y2, X2))
    ActiveChart.SeriesCollection(4).XValues = Sheets("OEB").Range(Cells(Y1, X1), Cells(Y2, X2))
    
    ActiveChart.SeriesCollection(4).Values = Sheets("OEB").Range(Cells(Y1 + 3, X1), Cells(Y2 + 3, X2))
    ActiveChart.SeriesCollection(3).Values = Sheets("OEB").Range(Cells(Y1 + 2, X1), Cells(Y2 + 2, X2))
    ActiveChart.SeriesCollection(2).Values = Sheets("OEB").Range(Cells(Y1 + 1, X1), Cells(Y2 + 1, X2))
    ActiveChart.SeriesCollection(1).Values = Sheets("OEB").Range(Cells(Y1 + 5, X1), Cells(Y2 + 5, X2))
    
    
   Next
    

'---------------OSW---------------
Sheets("OSW").Select
GRAF = Sheets("OSW").Range("AD" & 6).Value
X1 = Sheets("OSW").Range("AE" & 6).Value
X2 = Sheets("OSW").Range("AF" & 6).Value
Y1 = Sheets("OSW").Range("AG" & 6).Value
Y2 = Sheets("OSW").Range("AH" & 6).Value

ActiveSheet.ChartObjects("" & GRAF & "").Activate

ActiveChart.SeriesCollection(1).XValues = Sheets("OSW").Range(Cells(Y1, X1 - 2), Cells(Y2, X2 - 2))
ActiveChart.SeriesCollection(2).XValues = Sheets("OSW").Range(Cells(Y1, X1 - 2), Cells(Y2, X2 - 2))
ActiveChart.SeriesCollection(3).XValues = Sheets("OSW").Range(Cells(Y1, X1 - 2), Cells(Y2, X2 - 2))
    
ActiveChart.SeriesCollection(1).Values = Sheets("OSW").Range(Cells(Y1, X1 + 3), Cells(Y2, X2 + 3))
ActiveChart.SeriesCollection(2).Values = Sheets("OSW").Range(Cells(Y1, X1 + 4), Cells(Y2, X2 + 4))
ActiveChart.SeriesCollection(3).Values = Sheets("OSW").Range(Cells(Y1, X1 + 5), Cells(Y2, X2 + 5))

ActiveChart.SeriesCollection(1).Name = "=""ROFC"""
ActiveChart.SeriesCollection(2).Name = "=""DRNC"""
ActiveChart.SeriesCollection(3).Name = "=""PREC"""

Sheets("OSW").Select
GRAF = Sheets("OSW").Range("AD" & 7).Value
X1 = Sheets("OSW").Range("AE" & 7).Value
X2 = Sheets("OSW").Range("AF" & 7).Value
Y1 = Sheets("OSW").Range("AG" & 7).Value
Y2 = Sheets("OSW").Range("AH" & 7).Value

ActiveSheet.ChartObjects("" & GRAF & "").Activate

ActiveChart.SeriesCollection(1).XValues = Sheets("OSW").Range(Cells(Y1, X1 - 2), Cells(Y2, X2 - 2))
ActiveChart.SeriesCollection(2).XValues = Sheets("OSW").Range(Cells(Y1, X1 - 2), Cells(Y2, X2 - 2))
    
ActiveChart.SeriesCollection(1).Values = Sheets("OSW").Range(Cells(Y1, X1 + 1), Cells(Y2, X2 + 1))
ActiveChart.SeriesCollection(2).Values = Sheets("OSW").Range(Cells(Y1, X1 + 2), Cells(Y2, X2 + 2))


ActiveChart.SeriesCollection(1).Name = "=""SWTD"""
ActiveChart.SeriesCollection(2).Name = "=""SWXD"""

For OSW = 3 To 9

GRAF = Sheets("OSW").Range("AD" & OSW + 5).Value
X1 = Sheets("OSW").Range("AE" & OSW + 5).Value
X2 = Sheets("OSW").Range("AF" & OSW + 5).Value
Y1 = Sheets("OSW").Range("AG" & OSW + 5).Value
Y2 = Sheets("OSW").Range("AH" & OSW + 5).Value

   ActiveSheet.ChartObjects("" & GRAF & "").Activate
    
    ActiveChart.SeriesCollection(1).XValues = Sheets("OSW").Range(Cells(Y1, X1), Cells(Y2, X2))
    ActiveChart.SeriesCollection(2).XValues = Sheets("OSW").Range(Cells(Y1, X1), Cells(Y2, X2))
    ActiveChart.SeriesCollection(3).XValues = Sheets("OSW").Range(Cells(Y1, X1), Cells(Y2, X2))
    ActiveChart.SeriesCollection(4).XValues = Sheets("OSW").Range(Cells(Y1, X1), Cells(Y2, X2))
    
    ActiveChart.SeriesCollection(4).Values = Sheets("OSW").Range(Cells(Y1 + 3, X1), Cells(Y2 + 3, X2))
    ActiveChart.SeriesCollection(3).Values = Sheets("OSW").Range(Cells(Y1 + 2, X1), Cells(Y2 + 2, X2))
    ActiveChart.SeriesCollection(2).Values = Sheets("OSW").Range(Cells(Y1 + 1, X1), Cells(Y2 + 1, X2))
    ActiveChart.SeriesCollection(1).Values = Sheets("OSW").Range(Cells(Y1 + 5, X1), Cells(Y2 + 5, X2))
    
    
   Next
    

GRAF = Sheets("OSW").Range("AD" & OSW + 5).Value
X1 = Sheets("OSW").Range("AE" & OSW + 5).Value
X2 = Sheets("OSW").Range("AF" & OSW + 5).Value
Y1 = Sheets("OSW").Range("AG" & OSW + 5).Value
Y2 = Sheets("OSW").Range("AH" & OSW + 5).Value

ActiveSheet.ChartObjects("" & GRAF & "").Activate

ActiveChart.SeriesCollection(1).XValues = Sheets("OSW").Range(Cells(Y1, X1), Cells(Y2, X2))
ActiveChart.SeriesCollection(2).XValues = Sheets("OSW").Range(Cells(Y1, X1), Cells(Y2, X2))
ActiveChart.SeriesCollection(3).XValues = Sheets("OSW").Range(Cells(Y1, X1), Cells(Y2, X2))
    
ActiveChart.SeriesCollection(1).Values = Sheets("OSW").Range(Cells(Y1 + 2, X1), Cells(Y2 + 2, X2))
ActiveChart.SeriesCollection(2).Values = Sheets("OSW").Range(Cells(Y1 + 2, X1 + 75), Cells(Y2 + 2, X2 + 75))
ActiveChart.SeriesCollection(3).Values = Sheets("OSW").Range(Cells(Y1 + 2, X1 + 150), Cells(Y2 + 2, X2 + 150))

    
    
End Sub
