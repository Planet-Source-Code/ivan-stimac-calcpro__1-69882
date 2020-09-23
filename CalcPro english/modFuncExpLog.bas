Attribute VB_Name = "modFuncExpLog"
Option Explicit

'ln(z)
Public Function CLn(ByRef z As Complex) As Complex

    Dim r As Double, fi As Double
    
    r = Sqr(z.Im ^ 2 + z.Re ^ 2)
    fi = cArg(z) 'Atn(z.Im / z.Re)
    
    CLn.Re = Log(r)
    CLn.Im = fi
    
End Function

'log(z)
Public Function CLog(ByRef z As Complex, ByVal base As Double) As Complex
    
    If z.Im <> 0 Or z.Re < 0 Then
    
        Dim r As Double, fi As Double
        
        r = Sqr(z.Im ^ 2 + z.Re ^ 2)
        fi = cArg(z) 'Atn(z.Im / z.Re)
        
        CLog.Re = Log(r) / Log(base)
        CLog.Im = fi / Log(base)
        
    Else
    
        If z.Re = 0 Then
            lastErr = "-" & ERR_InfRes
            lastErrNum = ERR_InfResN
            Exit Function
        End If
        CLog.Re = Log(z.Re) / Log(base)
        
    End If
    
End Function

'exp(z)
Public Function CExp(ByRef z As Complex) As Complex

    CExp.Re = Exp(z.Re) * Cos(z.Im)
    CExp.Im = Exp(z.Re) * Sin(z.Im)
    
End Function

