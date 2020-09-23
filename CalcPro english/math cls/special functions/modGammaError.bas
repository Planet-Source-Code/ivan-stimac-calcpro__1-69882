Attribute VB_Name = "modGammaError"
Option Explicit
Private Const EulerMascheroni As Double = 0.577215664901533
Private Const mN As Integer = 100
Private Const FPMIN As Double = 1E-30

'/////////////////////////////////////////////////////////////////
'                       GAMMA FUNCTIONS
'/////////////////////////////////////////////////////////////////
Public Function functionGammaIncUpr(ByVal mX As Double, ByVal a As Double) As Double ' ide
    
    functionGammaIncUpr = functionGammaIncQ(mX, a) * functionGamma(mX)
    
End Function

Public Function functionGammaIncLwr(ByVal mX As Double, ByVal a As Double) As Double ' ide
    
    functionGammaIncLwr = functionGammaIncP(mX, a) * functionGamma(mX)
    
End Function

Public Function functionGammaIncP(ByVal mX As Double, ByVal a As Double) As Double

    Dim gammaRet As Double, gamLN As Double
    If mX > 0 And a > 0 Then
        
        If mX < (a + 1) Then
            gSer gammaRet, a, mX, gamLN
            functionGammaIncP = gammaRet
        Else
            gCf gammaRet, a, mX, gamLN
            functionGammaIncP = 1 - gammaRet
        End If
        
    End If
    
End Function


Public Function functionGammaIncQ(ByVal mX As Double, ByVal a As Double) As Double
    
    Dim gammaRet As Double, gamLN As Double
    If mX > 0 And a > 0 Then
        
        If mX < (a + 1) Then
            gSer gammaRet, a, mX, gamLN
            functionGammaIncQ = 1 - gammaRet
        Else
            gCf gammaRet, a, mX, gamLN
            functionGammaIncQ = gammaRet
        End If
        
    End If
    
End Function

Private Sub gSer(ByRef gamser As Double, ByVal a As Double, ByVal x As Double, ByRef gln As Double)
    
    Dim n As Integer
    Dim sum As Double, del As Double, ap As Double
    Const mN As Integer = 100
    Const EPS As Double = 0.0000003
    gln = functionGammaLN(a)
    
    If x < 0 Then
        
        lastErr = ERR_NegArg
        lastErrNum = ERR_NegArgN
        gamser = 0
        Exit Sub
        
    Else
        
        ap = a
        del = 1 / a
        sum = 1 / a
        For n = 1 To mN
            ap = ap + 1
            del = del * x / ap
            sum = sum + del
            If Abs(del) < Abs(sum) * EPS Then
                gamser = sum * Exp(-x + a * Log(x) - gln)
                Exit Sub
            End If
        Next n
        
    End If
    
End Sub

Private Sub gCf(ByRef gamser As Double, ByVal a As Double, ByVal x As Double, ByRef gln As Double)
    
    Const EPS As Double = 0.0000003
    Const FPMIN As Double = 1E-30
    Const mN  As Integer = 100
    Dim i As Integer
    Dim an As Double, b As Double, c As Double, d As Double, del As Double, h As Double
    
    gln = functionGammaLN(a)
    b = x + 1 - a
    c = 1 / FPMIN
    d = 1 / b
    h = d
    
    For i = 1 To mN
    
        an = -i * (i - a)
        b = b + 2
        d = an * d + b
        If Abs(d) < FPMIN Then d = FPMIN
        c = b + an / c
        If Abs(c) < FPMIN Then c = FPMIN
        d = 1 / d
        del = d * c
        h = h * del
        If Abs(del - 1) < EPS Then Exit For
        
    Next i
    
    gamser = Exp(-x + a * Log(x) - gln) * h
    
End Sub


Public Function functionGammaLN(ByVal mX As Double) As Double
    'On Error Resume Next
    If mX = 0 Then
        
        lastErr = ERR_InfRes
        lastErrNum = ERR_InfResN
        Exit Function
        
    ElseIf mX < 0 Then
        
        lastErr = ERR_NegArg
        lastErrNum = ERR_NegArgN
        Exit Function
        
    End If
    
    '
    Dim x As Double, y As Double, tmp As Double, ser As Double
    Dim arr(5) As Double
    Dim j As Integer
    arr(0) = 76.1800917294715
    arr(1) = -86.5053203294168
    arr(2) = 24.0140982408309
    arr(3) = -1.23173957245015
    arr(4) = 1.20865097386618E-03
    arr(5) = -5.395239384953E-06
    
    y = mX
    x = mX
    tmp = x + 5.5
    tmp = tmp - (x + 0.5) * Log(tmp)
    ser = 1.00000000019001
    For j = 0 To 5
        y = y + 1
'        If y <> 0 Then
            ser = ser + arr(j) / y
'        Else
'            isResInf = True
'            Exit Function
'        End If
    Next j
    
    functionGammaLN = -tmp + Log(2.506628274631 * ser / mX)

    '3.1781
End Function
'Public Function functionGamma(ByVal mX As Double) As Double
'    If mX > 0 Then
'        functionGamma = Exp(functionGammaLN(mX))
'    'za nula i sve negativne cjelobrojne vrijednosti
'    ElseIf mX = 0 Or mX = Fix(mX) Then
'        lastErr = ERR_InfRes
'        lastErrNum = ERR_InfResN
'        Exit Function
'    ElseIf mX < 0 Then
'        'functionGamma = functionGammaNeg(mX)
'
'    End If
'
'
'End Function

'Lanczos approximation
'The Lanczos approximation offers a relatively simple way to
'calculate the gamma function of a complex argument to any
'desired precision.
'The precision is determined by a free constant g and a series
'whose coefficients depend on g and may be computed in advanced
'if a fixed precision is desired. The following implementation
'uses the same coefficients as the GNU Scientific Library's gamma function.
'The typical relative error is about 10-15 (nearly full double precision accuracy).

'Lanczos  series is only valid for numbers in the right half complex plane, so the reflection formula for the gamma function is employed for negative arguments (and also arguments in (0, 0.5), since this should increase accuracy near the singularity at zero.)

Public Function functionGamma(ByVal mX As Double)
    
    If mX = 0 Or (mX = Fix(mX) And mX < 0) Then
        
        lastErr = ERR_InfRes
        lastErrNum = ERR_InfResN
        Exit Function
        
    End If
    '
    Dim mArr(8) As Double, x As Double, t As Double
    Dim i As Integer
    Const cNum As Integer = 7   'const g
    
    mArr(0) = 0.99999999999981
    mArr(1) = 676.520368121885
    mArr(2) = -1259.1392167224
    mArr(3) = 771.323428777653
    mArr(4) = -176.615029162141
    mArr(5) = 12.5073432786869
    mArr(6) = -0.13857109526572
    mArr(7) = 9.98436957801957E-06
    mArr(8) = 1.50563273514931E-07

    If mX < 0.5 Then
        
        functionGamma = PI / (Sin(PI * mX) * functionGamma(1 - mX))
        
    Else
    
        mX = mX - 1
        x = mArr(0)
        
        For i = 1 To cNum + 1
        
            x = x + mArr(i) / (mX + i)
            
        Next i
        
        t = mX + cNum + 0.5
        functionGamma = Sqr(2 * PI) * t ^ (mX + 0.5) * Exp(-t) * x
        
    End If
End Function

'Private Function functionGammaNeg(ByVal mX As Double) As Double ' ide
'    Dim p(6) As Double, sum1 As Double ', subRes As Double
'    Dim mRes As Double
'    Dim n As Integer
'
'
'    p(0) = 1.00000000019001
'    p(1) = 76.1800917294715
'    p(2) = -86.5053203294168
'    p(3) = 24.0140982408309
'    p(4) = -1.23173957245015
'    p(5) = 1.20865097386618E-03
'    p(6) = -5.395239384953E-06
'
'    sum1 = p(0)
'    For n = 1 To 6
'         sum1 = sum1 + p(n) / (n + mX)
'    Next n
'    ' MsgBox Exp(-mx - 5.5)
'    mRes = sum1 * Sqr(2 * PI) / mX * (mX + 5.5) ^ (mX + 0.5) * Exp(-mX - 5.5)
'
'    functionGammaNeg = mRes
'End Function

'/////////////////////////////////////////////////////////////////
'                       ERROR FUNCTIONS
'/////////////////////////////////////////////////////////////////
'error function
Public Function functionErf(ByVal mX As Double) As Double

    If mX < 0 Then
        
        functionErf = -functionGammaIncP(mX * mX, 0.5)
        
    Else
        
        functionErf = functionGammaIncP(mX * mX, 0.5)
        
    End If
    
End Function
'complementary error function
Public Function functionErfC(ByVal mX As Double) As Double

    If mX <= 0 Then
    
        functionErfC = 1 + functionGammaIncP(mX * mX, 0.5)
        
    Else
    
        functionErfC = functionGammaIncQ(mX * mX, 0.5)
        
    End If
    
End Function
'Scaled complementary error function
Public Function functionErfCX(ByVal mX As Double) As Double

    functionErfCX = Exp(mX ^ 2) * functionErfC(mX)
    
End Function

'/////////////////////////////////////////////////////////////////
'                       BETA FUNCTION
'/////////////////////////////////////////////////////////////////

Public Function functionBeta(ByVal a As Double, ByVal b As Double) As Double
    'provjera argumenata
    If a < 0 Or b < 0 Then
        
        lastErr = ERR_NegArg
        lastErrNum = ERR_NegArgN
        Exit Function
        
    ElseIf functionGamma(a + b) = 0 Then

        lastErr = ERR_InfRes
        lastErrNum = ERR_InfResN
        Exit Function
        
    End If
    
    functionBeta = functionGamma(a) * functionGamma(b) / functionGamma(a + b)
    
End Function


Public Function functionBetaLn(ByVal a As Double, ByVal b As Double) As Double
    
    If a < 0 Or b < 0 Then
        
        lastErr = ERR_NegArg
        lastErrNum = ERR_NegArgN
        Exit Function
        
    ElseIf functionBeta(a, b) = 0 Then
        
        lastErr = ERR_InfRes
        lastErrNum = ERR_InfResN
        Exit Function
        
    Else
    
        functionBetaLn = Log(functionBeta(a, b))
        
    End If
    
End Function


Public Function functionBetaInc(ByVal mX As Double, ByVal a As Double, ByVal b As Double) As Double
    
    Const EPS As Double = 0.0000003
    Dim bt As Double
    
    If mX < 0 Or mX > 1 Then
    
        lastErr = ERR_BetaIncErr & " Invalid input arguments"
        lastErrNum = ERR_BetaIncErrN
        Exit Function
        
    Else
        
        If mX = 0 Or mX = 1 Then
            bt = 0
        Else
            bt = Exp(functionGammaLN(a + b) - functionGammaLN(a) - functionGammaLN(b) + a * Log(mX) + b * Log(1 - mX))
        End If
        
        If a = b Then
        
            If mX = 0.5 Or mX = 1 Then
                functionBetaInc = mX
                Exit Function
            End If
            
        End If

        If mX < (a + 1) / (a + b + 2) Then
            
            If a = 0 Then
                lastErr = ERR_InfRes
                lastErrNum = ERR_InfResN
                Exit Function
            End If
            functionBetaInc = bt * getBetaCf(a, b, mX) / a
            
        Else
            
            If b = 0 Then
                lastErr = ERR_InfRes
                lastErrNum = ERR_InfResN
                Exit Function
            End If
           
            functionBetaInc = 1 - bt * getBetaCf(b, a, 1 - mX) / b
        End If
    End If
    
End Function
Private Function getBetaCf(ByVal a As Double, ByVal b As Double, ByVal mX As Double) As Double
    
    Const EPS As Double = 0.0000003
    Dim m As Integer, m2 As Integer
    Dim aa As Double, c As Double, d As Double, del As Double, h As Double
    Dim qab As Double, qam As Double, qap As Double
    
    qab = a + b
    qap = a + 1
    qam = a - 1
    c = 1
    d = 1 - qab * mX / qap
    If Abs(d) < FPMIN Then d = FPMIN
    d = 1 / d
    h = d
    
    For m = 1 To mN
        
        m2 = 2 * m
        aa = m * (b - m) * mX / ((qam + m2) * (a + m2))
        d = 1 + aa * d
        If Abs(d) < FPMIN Then d = FPMIN
        c = 1 + aa / c
        If Abs(c) < FPMIN Then c = FPMIN
        d = 1 / d
        h = h * d * c
        aa = -(a + m) * (qab + m) * mX / ((a * m2) * (qap + m2))
        d = 1 + aa * d
        If Abs(d) < FPMIN Then d = FPMIN
        c = 1 + aa / c
        If Abs(c) < FPMIN Then c = FPMIN
        d = 1 / d
        del = d * c
        h = h * del
        If Abs(del - 1) < EPS Then Exit For
        
    Next m
    
    If m > mN Then

        lastErr = ERR_BetaErr
        lastErrNum = ERR_BetaErrN
        Exit Function
        
    Else
    
        getBetaCf = h
        
    End If

End Function
