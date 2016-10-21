Attribute VB_Name = "Módulo32"
Sub Macro37()

Dim par_n As Integer
Dim res_matrix(1 To 100, 1 To 100) As Double

RES_ABS1 = Range("P" & 2).Value
par_n = Range("O" & 2).Value

Dim p() As Double
Dim r() As Double
ReDim p(par_n)
ReDim r(par_n)

For x = 1 To par_n

    p(x) = Range("G" & x + 1).Value
    r(x) = Range("H" & x + 1).Value

Next

res_befo = Range("P" & 2).Value
    
    
For y = 1 To par_n
h = 1
For h = 1 To par_n
Range("J" & h + 1).Select
ActiveCell.FormulaR1C1 = p(h)
Next
    
        For Z = 1 To 100
        Range("J" & y + 1).Select
        ActiveCell.FormulaR1C1 = p(y) - r(y) + (r(y) * 2) * (Z / 100)
        Calculate
        res_matrix(1, Z) = p(y) - r(y) + (r(y) * 2) * (Z / 100)
        res_matrix(2, Z) = Abs(Range("P" & 2).Value)
        
        Range(Cells(Z + 1, 20), Cells(Z + 1, 20)).Select
        ActiveCell.FormulaR1C1 = res_matrix(1, Z)
        Range(Cells(Z + 1, 21), Cells(Z + 1, 21)).Select
        ActiveCell.FormulaR1C1 = res_matrix(2, Z)
        'rever o proximo passo...
        If res_matrix(2, Z) < res_befo Then
        p(y) = res_matrix(1, Z)
        res_befo = res_matrix(2, Z)
        End If
        
        
        Next
        
        
        
    Next



x = p(2)



Do While RES_ABS1 > 2

Calculate
RES_ABS2 = Range("D" & 14).Value

'GoTo TCHAU

If RES_ABS2 < RES_ABS1 Then


p1 = Range("J" & 2).Value
p2 = Range("J" & 3).Value
p3 = Range("J" & 4).Value

R1 = Range("H" & 2).Value
R2 = Range("H" & 3).Value
R3 = Range("H" & 4).Value

Range("k2").Select
ActiveCell.FormulaR1C1 = p1
Range("k3").Select
ActiveCell.FormulaR1C1 = p2
Range("k4").Select
ActiveCell.FormulaR1C1 = p3


Range("k5").Select
ActiveCell.FormulaR1C1 = RES_ABS2

'R1 = R1 * RES_ABS2 / RES_ABS1
'R2 = R2 * RES_ABS2 / RES_ABS1
'R3 = R3 * RES_ABS2 / RES_ABS1

Range("H2").Select
ActiveCell.FormulaR1C1 = R1
Range("H3").Select
ActiveCell.FormulaR1C1 = R2
Range("H4").Select
ActiveCell.FormulaR1C1 = R3


Range("G2").Select
ActiveCell.FormulaR1C1 = p1
Range("G3").Select
ActiveCell.FormulaR1C1 = p2
Range("G4").Select
ActiveCell.FormulaR1C1 = p3


RES_ABS1 = RES_ABS2
End If

TCHAU:

Range("L2").Select
ActiveCell.FormulaR1C1 = K

K = K + 1

Loop

End Sub

Sub comb()

Dim a, b, c As Double
Dim Ra, Rb, Rc As Double
' Number of combinations will be (n+1)^npar
' npar = number of parameters (a,b,c,d,e...)
' ni   = is the botton range for the iTh parameter
' Ri   = is the above range of the iTh parameter


ai = 1
bi = 3
ci = 10
di = 12
ei = 0.3

a = ai
b = bi
c = ci
d = di
e = ei

Ra = 2
Rb = 5
Rc = 2
Rd = 13
Re = 0.5

n = 10
l = 1
For xa = 1 To n + 1

    b = bi
    For xb = 1 To n + 1
    
        c = ci
        For xc = 1 To n + 1
        
            d = di
            For xd = 1 To n + 1
                
                e = ei
                For xe = 1 To n + 1
                  
                    Range("D" & l & "").Value = a
                    Range("E" & l & "").Value = b
                    Range("F" & l & "").Value = c
                    Range("G" & l & "").Value = d
                    Range("H" & l & "").Value = e
                    
                    l = l + 1
                    
                e = e + (Re / n)
                Next
                                
            d = d + (Rd / n)
            Next
            
        c = c + (Rc / n)
        Next
    
    b = b + (Rb / n)
    Next
    
a = a + (Ra / n)
Next

End Sub

Sub abre()

'Abre um determinado arquivo pelo cmd.exe
'/c para fechar no final e /k para manter aberto no final
'START comando do cmd para abrir um arquivo
'/MIN minimizado /MAX maximizado

    'Shell "C:\Windows\System32\cmd.exe """ & "/c START /D c:\Murilo\MESTRADO\Simulacao\Batch_DSSAT\Sequence\ /MAX BATCH_DSSAT.bat " & "", vbNormalFocus
     Shell "C:\Windows\System32\cmd.exe """ & "/k cd C:\Murilo\DOUTORADO\WUR\SWAP_Sugarcanev2\SWAP_Sugarcanev2\SWAP_Sugarcanev2" & "" & "k/ SWAP_Sugarcanev2" & "", vbNormalFocus
           
           
           
End Sub

