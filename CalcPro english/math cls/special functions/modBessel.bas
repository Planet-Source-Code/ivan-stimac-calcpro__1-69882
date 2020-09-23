Attribute VB_Name = "modBessel"
Option Explicit

Private Function functionBesselJ0(ByVal mX As Double) As Double
    Dim ax As Double, z As Double, xx As Double
    Dim y As Double, ans As Double, ans1 As Double, ans2 As Double
    ax = Abs(mX)
    If ax < 8 Then
        y = mX * mX
        ans1 = -11214424.18 + y * (77392.33017 + y * (-184.9052456))
        ans1 = ans1 * y + 651619640.7
        ans1 = ans1 * y - 13362590354#
        ans1 = ans1 * y + 57568490574#
        '
        ans2 = 9494680.718 + y * (59272.64853 + y * (267.8532712 + y * 1))
        ans2 = 57568490411# + y * (1029532985 + y * (ans2))
        ans = ans1 / ans2
    Else
        z = 8 / ax
        y = z * z
        xx = ax - 0.785398164
        ans1 = (0.00002734510407 + y * (-0.000002073370639 + y * 2.093887211E-07))
        ans1 = ans1 * y - 0.001098628627
        ans1 = ans1 * y + 1
        
        ans2 = -0.000006911147651 + y * (7.621095161E-07 - y * 9.34945152E-08)
        ans1 = ans1 * y + 0.0001430488765
        ans1 = ans1 * y - 0.01562499995
    
        ans = Sqr(0.636619772 / ax) * (Cos(xx) * ans1 - z * Sin(xx) * ans2)
    End If
    functionBesselJ0 = ans
End Function


Private Function functionBesselY0(ByVal mX As Double) As Double
    Dim z As Double, xx As Double, y As Double, ans As Double
    Dim ans1 As Double, ans2 As Double
    If mX < 8 Then
        y = mX * mX
        ans1 = 10879881.29 + y * (-86327.92757 + y * 228.4622733)
        ans1 = ans1 * y - 512359803.6
        ans1 = ans1 * y + 7062834065#
        ans1 = ans1 * y - 2957821389#
        
        ans2 = 7189466.438 + y * (47447.2647 + y * (226.1030244 + y))
        ans2 = ans2 * y + 745249964.8
        ans2 = ans2 * y + 40076544269#
        
        ans = (ans1 / ans2) + 0.636619772 * functionBesselJ0(mX) * Log(mX)
    Else
        z = 8 / mX
        y = z * z
        xx = mX - 0.785398164
        
        ans1 = 0.00002734510407 + y * (-0.000002073370639 + y * 2.093887211E-07)
        ans1 = ans1 * y - 0.001098628627
        ans1 = ans1 * y + 1
        '
        ans2 = -0.000006911147651 + y * (7.621095161E-07 + y * (-9.34945152E-08))
        ans2 = ans2 * y + 0.0001430488765
        ans2 = ans2 * y - 0.01562499995
        
        ans = Sqr(0.636619772 / mX) * (Sin(xx) * ans1 + z * Cos(xx) * ans2)
    End If
    functionBesselY0 = ans
End Function

Private Function functionBesselJ1(ByVal mX As Double) As Double
    Dim xx As Double, z As Double, ax As Double, y As Double
    Dim ans As Double, ans1 As Double, ans2 As Double
    ax = Abs(mX)
    If ax < 8 Then
        y = mX * mX
        
        ans1 = -2972611.439 + y * (15704.4826 + y * (-30.16036606))
        ans1 = ans1 * y + 242396853.1
        ans1 = ans1 * y - 7895059235#
        ans1 = ans1 * y + 72362614232#
        ans1 = ans1 * mX
        
        ans2 = 18583304.74 + y * (99447.43394 + y * (376.9991397 + y))
        ans2 = ans2 * y + 2300535178#
        ans2 = ans2 * y + 144725228442#
        
        ans = ans1 / ans2
    Else
        z = 8 / ax
        y = z * z
        xx = ax - 2.356194491
        ans1 = -0.00003516396496 + y * (0.000002457520174 + y * (-0.000000240337019))
        ans1 = y * ans1 + 0.00183105
        ans1 = y * ans1 + 1
        
        ans2 = 0.000008449199096 + y * (-0.00000088228987 + y * 0.000000105787412)
        ans2 = ans2 * y - 0.0002002690873
        ans2 = ans2 * y + 0.04687499995
        
        ans = Sqr(0.636619772 / ax) * (Cos(xx) * ans1 - z * Sin(xx) * ans2)
    End If
    functionBesselJ1 = ans
End Function

Private Function functionBesselY1(ByVal mX As Double) As Double
    Dim z As Double, xx As Double, y As Double, ans As Double
    Dim ans1 As Double, ans2 As Double
    If mX < 8 Then
        y = mX * mX
        ans1 = -51534381390# + y * (734926455.1 + y * (-4237922.726 + y * 8511.937935))
        ans1 = ans1 * y + 1275274390000#
        ans1 = ans1 * y - 4900604943000#
        ans1 = ans1 * mX
        
        ans2 = 3733650367# + y * (22459040.02 + y * (102042.605 + y * (354.9632885 + y)))
        ans2 = ans2 * y + 424441966400#
        ans2 = ans2 * y + 24995805700000#
        ans = ans1 / ans2 + 0.636619772 * (functionBesselJ1(mX) * Log(mX) - 1 / mX)
    Else
        z = 8 / mX
        y = z * z
        xx = mX - 2.356194491
        ans1 = -0.00003516396496 + y * (0.000002457520174 + y * (-0.000000240337019))
        ans1 = ans1 * y + 0.00183105
        ans1 = ans1 * y + 1
        
        ans2 = 0.000008449199096 + y * (-0.00000088228987 + y * 0.000000105787412)
        ans2 = ans2 * y - -0.0002002690873
        ans2 = ans2 * y + 0.04687499995
        ans = Sqr(0.636619772 / mX) * (Sin(xx) * ans1 + z * Cos(xx) * ans2)
    End If
    functionBesselY1 = ans
End Function


Public Function functionBesselY(ByVal mX As Double, ByVal n As Integer) As Double
    Dim j As Integer
    Dim by As Double, bym As Double, byp As Double, tox As Double
    If n < 0 Then
        'error:invalid n
    ElseIf n = 0 Then
        functionBesselY = functionBesselY0(mX)
    ElseIf n = 1 Then
        functionBesselY = functionBesselY1(mX)
    Else
        tox = 2 / mX
        by = functionBesselY1(mX)
        bym = functionBesselY0(mX)
        For j = 1 To n - 1
            byp = j * tox * by - bym
            bym = by
            by = byp
        Next j
        functionBesselY = by
    End If
End Function

Public Function functionBesselJ(ByVal mX As Double, ByVal n As Integer) As Double
    Const ACC As Integer = 40
    Const BIGNO As Double = 10000000000#
    Const BIGNI  As Double = 0.0000000001
    Dim j As Integer, jsum As Integer, m As Integer
    Dim ax As Double, bj As Double, bjm As Double, bjp As Double, sum As Double
    Dim tox As Double, ans As Double
    
    If n < 2 Then
        'error
        lastErr = ERR_BesselErr & vbCrLf & "Second param must be >= 2"
        lastErrNum = ERR_BesselErrN
        Exit Function
    'ElseIf n = 0 Then
        'functionBesselJ = functionBesselY0(mX)
    'ElseIf n = 1 Then
        'functionBesselJ = functionBesselY1(mX)
    Else
        ax = Abs(mX)
        If ax = 0 Then
            functionBesselJ = 0
            Exit Function
        ElseIf (ax > n) Then
            tox = 2 / ax
            bjm = functionBesselJ0(ax)
            bj = functionBesselJ1(ax)
            For j = 1 To n - 1
                bjp = j * tox * bj - bjm
                bjm = bj
                bj = bjp
            Next j
            ans = bj
        Else
            tox = 2 / ax
            m = 2 * ((n + Int(Sqr(ACC * n))) / 2)
            jsum = 0
            bjp = 0
            ans = 0
            sum = 0
            bj = 1
            For j = m To 1 Step -1
                bjm = j * tox * bj - bjp
                bjp = bj
                bj = bjm
                If Abs(bj) > BIGNO Then
                    bj = bj * BIGNI
                    bjp = bjp * BIGNI
                    ans = ans * BIGNI
                    sum = sum * BIGNI
                End If
                If jsum Then sum = sum + bj
                jsum = Not jsum
                If j = n Then ans = bjp
            Next j
            sum = 2 * sum - bj
            ans = ans / sum
        End If
        If mX < 0 And (n And 1) Then
            functionBesselJ = -ans
        Else
            functionBesselJ = ans
        End If
    End If
End Function


'-----------------------------------------------------------
Private Function functionBesselI0(ByVal mX As Double) As Double
    Dim ax As Double, ans As Double, y As Double
    
    ax = Abs(mX)
    If ax < 3.75 Then
        y = mX / 3.75
        y = y * y
        ans = 1.2067492 + y * (0.2659732 + y * (0.0360768 + y * 0.0045813))
        ans = ans * y + 3.0899424
        ans = ans * y + 3.5156229
        ans = ans * y + 1
    Else
        y = 3.75 / ax
        ans = -0.02057706 + y * (0.02635537 + y * (-0.01647633 + y * 0.00392377))
        ans = y * ans + 0.00916281
        ans = y * ans - 0.00157565
        ans = y * ans + 0.00225319
        ans = y * ans + 0.01328592
        ans = y * ans + 0.39894228
        ans = ans * (Exp(ax) / Sqr(ax))
    End If
    functionBesselI0 = ans
End Function
'1
Private Function functionBesselI1(ByVal mX As Double) As Double
    Dim ax As Double, ans As Double, y As Double
    ax = Abs(mX)
    If ax < 3.75 Then
        y = mX / 3.75
        y = y * y
        ans = 0.15084934 + y * (0.02658733 + y * (0.00301532 + y * 0.00032411))
        ans = ans * y + 0.51498869
        ans = ans * y + 0.87890594
        ans = ans * y + 0.5
        ans = ax * ans
    Else
        y = 3.75 / ax
        
        ans = 0.02282967 + y * (-0.02895312 + y * (0.01787654 - y * 0.00420059))
        ans = -0.00362018 + y * (0.00163801 + y * (-0.01031555 + y * ans))
        ans = 0.39894228 + y * (-0.03988024 + y * ans)
        ans = ans * (Exp(ax) / Sqr(ax))
    End If
    functionBesselI1 = ans
End Function

Private Function functionBesselK0(ByVal mX As Double) As Double
    Dim ans As Double, y As Double
    If mX <= 2 Then
        y = mX * mX / 4
        ans = 0.0348859 + y * (0.00262698 + y * (0.0001075 + y * 0.0000074))
        ans = ans * y + 0.23069756
        ans = ans * y + 0.4227842
        ans = ans * y - 0.57721566
        ans = ans + (-Log(mX / 2) * functionBesselI0(mX))
    Else
        y = 2 / mX
        ans = -0.01062446 + y * (0.00587872 + y * (-0.0025154 + y * 0.00053208))
        ans = ans * y + 0.02189568
        ans = ans * y - 0.07832358
        ans = ans * y + 1.25331414
        ans = ans * (Exp(-mX) / Sqr(mX))
    End If
    functionBesselK0 = ans
End Function
Private Function functionBesselK1(ByVal mX As Double) As Double
    Dim ans As Double, y As Double
    If mX <= 2 Then
        y = mX * mX / 4
        ans = -0.01919402 + y * (-0.00110404 + y * (-0.00004686))
        ans = 0.15443144 + y * (-0.67278579 + y * (-0.18156897 + y * ans))
        ans = (Log(mX / 2) * functionBesselI1(mX)) + (1 / mX) * (1 + y * ans)
    Else
        y = 2 / mX
        ans = 0.01504268 + y * (-0.00780353 + y * (0.00325614 + y * (-0.00068245)))
        ans = 1.25331414 + y * (0.23498619 + y * (-0.0365562 + y * ans))
        ans = (Exp(-mX) / Sqr(mX)) * (ans)
    End If
    functionBesselK1 = ans
End Function

Public Function functionBesselK(ByVal mX As Double, ByVal n As Integer) As Double
    Dim j As Integer
    Dim bk As Double, bkm As Double, bkp As Double, tox As Double
    
    If n < 0 Then
        'error
    ElseIf n = 0 Then
        functionBesselK = functionBesselK0(mX)
    ElseIf n = 1 Then
        functionBesselK = functionBesselK1(mX)
    Else
        tox = 2 / mX
        bkm = functionBesselK0(mX)
        bk = functionBesselK1(mX)
        For j = 1 To n - 1
            bkp = bkm + j * tox * bk
            bkm = bk
            bk = bkp
        Next j
        functionBesselK = bk
    End If
End Function

Public Function functionBesselI(ByVal mX As Double, ByVal n As Integer) As Double
    Const ACC As Integer = 40
    Const BIGNO As Double = 10000000000#
    Const BIGNI As Double = 0.0000000001
    Dim j As Integer
    Dim bi As Double, bim As Double, bip As Double, tox As Double, ans As Double
    
    If n < 0 Then
        'error
    ElseIf n = 0 Then
        functionBesselI = functionBesselI0(mX)
    ElseIf n = 1 Then
        functionBesselI = functionBesselI1(mX)
    Else
        tox = 2 / Abs(mX)
        bip = 0
        ans = 0
        bi = 1
        For j = 2 * (n + Int(Sqr(ACC * n))) To 1 Step -1
            bim = bip + j * bi * tox
            bip = bi
            bi = bim
            If Abs(bi) > BIGNO Then
                ans = ans * BIGNO
                bi = bi * BIGNO
                bip = bip * BIGNO
            End If
            If j = n Then ans = bip
        Next j
        ans = ans * functionBesselI0(mX) / bi
        If mX < 0 And (n And 1) Then
            functionBesselI = -ans
        Else
            functionBesselI = ans
        End If
    End If
End Function

'Returns the Bessel functions rj & ry
'and their derivatives rjp & ryp
Private Sub getBesselJY(ByVal x As Double, ByVal xnu As Double, ByRef rj As Double, _
            ByRef ry As Double, ByRef rjp As Double, ByRef ryp As Double)
    Dim i As Integer, isign As Integer, l As Integer, nl As Integer
    Dim a As Double, b As Double, br As Double, bi As Double, c As Double, cr As Double, ci As Double
    Dim d As Double, del As Double, del1 As Double, den As Double, di As Double, dr As Double, dlr As Double, dli As Double
    Dim e As Double, f As Double, fact As Double, fact2 As Double
    Dim fact3 As Double, ff As Double, gam As Double, gam1 As Double, gam2 As Double, gammi As Double, gampl As Double
    Dim h As Double, p As Double, pimu As Double, pimu2 As Double, q As Double, r As Double, rjl As Double
    Dim rjl1 As Double, rjmu As Double, rjp1 As Double, rjpl As Double, rjtemp As Double
    Dim ry1 As Double, rymu As Double, rymup As Double, rytemp As Double, sum As Double, sum1 As Double
    Dim temp As Double, w As Double, x2 As Double, xi As Double, xi2 As Double, xmu As Double, xmu2 As Double
    
    Const EPS As Double = 0.0000000001
    Const FPMIN As Double = 1E-30
    Const MAXIT As Integer = 10000
    Const XMIN As Integer = 2
    
    If x <= 0 Or xnu < 0 Then
        'error:"bad arguments in bessjy"
        Exit Sub
    End If
    
    If x < XMIN Then
        nl = Int(xnu + 0.5)
    Else
        If Int(xnu - x + 1.5) < 0 Then
            nl = 0
        Else
            nl = Int(xnu - x + 1.5)
        End If
    End If
    
    xmu = xnu - nl
    xmu2 = xmu * xmu
    xi = 1 / x
    xi2 = 2 * xi
    w = xi2 / PI
    isign = 1
    h = xnu * xi
    
    If h < FPMIN Then h = FPMIN
    b = xi2 * xnu
    d = 0
    c = h
    
    For i = 1 To MAXIT
        b = b + xi2
        d = b - d
        If Abs(d) < FPMIN Then d = FPMIN
        c = b - 1 / c
        If Abs(c) < FPMIN Then c = FPMIN
        d = 1 / d
        del = c * d
        h = del * h
        If d < 0 Then isign = -isign
        If Abs(del - 1) < EPS Then Exit For
    Next i
    
    If (i > MAXIT) Then
        'erro: "x too large in bessjy; try asymptotic expansion"
        'MsgBox "x too large in bessjy; try asymptotic expansion"
        lastErr = ERR_BesselErr
        lastErrNum = ERR_BesselErrN
        Exit Sub
    End If
    rjl = isign * FPMIN
    rjpl = h * rjl
    rjl1 = rjl
    rjp1 = rjpl
    fact = xnu * xi
    
    For l = nl To 1 Step -1
        rjtemp = fact * rjl + rjpl
        fact = fact - xi
        rjpl = fact * rjtemp - rjl
        rjl = rjtemp
    Next l
    
    If rjl = 0 Then rjl = EPS
    f = rjpl / rjl
    
    If (x < XMIN) Then
        x2 = 0.5 * x
        pimu = PI * xmu
        If Abs(pimu) < EPS Then
            fact = 1
        Else
            fact = pimu / Sin(pimu)
        End If
        d = -Log(x2)
        e = xmu * d
        
        If Abs(e) < EPS Then
            fact2 = 1
        Else
            'sinh(e)=(Exp(e) - Exp(-e)) / 2
            fact2 = (Exp(e) - Exp(-e)) / 2 / e
        End If
        
        getBeschb xmu, gam1, gam2, gampl, gammi
        
                                    'cosh(e)
        ff = 2 / PI * fact * (gam1 * (Exp(e) + Exp(-e)) / 2 + gam2 * fact2 * d)
        e = Exp(e)
        p = e / (gampl * PI)
        q = 1 / (e * PI * gammi)
        pimu2 = 0.5 * pimu
        
        'fact3 = (fabs(pimu2) < EPS ? 1.0 : sin(pimu2)/pimu2);
        If Abs(pimu2) < EPS Then
            fact3 = 1
        Else
            fact3 = Sin(pimu2) / pimu2
        End If
        r = PI * pimu2 * fact3 * fact3
        c = 1
        d = -x2 * x2
        sum = ff + r * q
        sum1 = p
        
        For i = 1 To MAXIT
            ff = (i * ff + p + q) / (i * i - xmu2)
            c = c * (d / i)
            p = p / (i - xmu)
            q = q / (i + xmu)
            del = c * (ff + r * q)
            sum = sum + del
            del1 = c * p - i * del
            sum1 = sum1 + del1
            If Abs(del) < (1 + Abs(sum)) * EPS Then Exit For
        Next i
        If (i > MAXIT) Then
            'error:nrerror("bessy series failed to converge");
            lastErr = ERR_BesselErr
            lastErrNum = ERR_BesselErrN
            Exit Sub
        End If
        rymu = -sum
        ry1 = -sum1 * xi2
        rymup = xmu * xi * rymu - ry1
        rjmu = w / (rymup - f * rymu)
    Else
        a = 0.25 - xmu2
        p = -0.5 * xi
        q = 1
        br = 2 * x
        bi = 2
        fact = a * xi / (p * p + q * q)
        cr = br + q * fact
        ci = bi + p * fact
        den = br * br + bi * bi
        dr = br / den
        di = -bi / den
        dlr = cr * dr - ci * di
        dli = cr * di + ci * dr
        temp = p * dlr - q * dli
        q = p * dli + q * dlr
        p = temp
        
        For i = 2 To MAXIT
            a = a + 2 * (i - 1)
            bi = bi + 2
            dr = a * dr + br
            di = a * di + bi
            If Abs(dr) + Abs(di) < FPMIN Then dr = FPMIN
            fact = a / (cr * cr + ci * ci)
            cr = br + cr * fact
            ci = bi - ci * fact
            If Abs(cr) + Abs(ci) < FPMIN Then cr = FPMIN
            den = dr * dr + di * di
            dr = dr / den
            di = di / -den
            dlr = cr * dr - ci * di
            dli = cr * di + ci * dr
            temp = p * dlr - q * dli
            q = p * dli + q * dlr
            p = temp
            If Abs(dlr - 1) + Abs(dli) < EPS Then Exit For
        Next i
        If (i > MAXIT) Then
            'error:nrerror("cf2 failed in bessjy");
            lastErr = ERR_BesselErr
            lastErrNum = ERR_BesselErrN
            Exit Sub
        End If
        gam = (p - f) / q
        rjmu = Sqr(w / ((p - f) * gam + q))
        'SIGN(a,b) ((b) >= 0.0 ? fabs(a) : -fabs(a))
        If rjl >= 0 Then
            rjmu = Abs(rjmu)
        Else
            rjmu = -Abs(rjmu)
        End If
        '
        'rjmu = SIGN(rjmu, rjl)
        rymu = rjmu * gam
        rymup = rymu * (p + q / gam)
        ry1 = xmu * xi * rymu - rymup
    End If
    fact = rjmu / rjl
    rj = rjl1 * fact
    rjp = rjp1 * fact
    
    For i = 1 To nl
        rytemp = (xmu + i) * xi2 * ry1 - rymu
        rymu = ry1
        ry1 = rytemp
    Next i
    ry = rymu
    ryp = xnu * xi * rymu - ry1

End Sub

Private Sub getBeschb(ByVal x As Double, ByRef gam1 As Double, ByRef gam2 As Double, ByRef gampl As Double, ByRef gammi As Double)
    Const NUSE1 As Integer = 5
    Const NUSE2 As Integer = 5
    Dim xx As Double
    Dim c1(6) As Double, c2(7) As Double
    
    c1(0) = -1.14202268037117
    c1(1) = 6.5165112670737E-03
    c1(2) = 3.087090173086E-04
    c1(3) = -3.4706269649E-06
    c1(4) = 6.9437664E-09
    c1(5) = 3.67795E-11
    c1(6) = -1.356E-13
    
    c2(0) = 1.8437405873009
    c2(1) = -7.68528408447867E-02
    c2(2) = 1.2719271366546E-03
    c2(3) = -4.9717367042E-06
    c2(4) = -3.31261198E-08
    c2(5) = 2.423096E-10
    c2(6) = -1.702E-13
    c2(7) = -1.49E-15
    
    xx = 8 * x * x - 1
    gam1 = getChebev(-1, 1, c1, NUSE1, xx)
    gam2 = getChebev(-1, 1, c2, NUSE2, xx)
    gampl = gam2 - x * (gam1)
    gammi = gam2 + x * (gam1)
End Sub
'Chebyshev evaluation
Private Function getChebev(ByVal a As Double, ByVal b As Double, ByRef c() As Double, ByVal m As Integer, ByVal x As Double) As Double
    Dim d As Double, dd As Double, sv As Double, y As Double, y2 As Double
    Dim j As Integer
    d = 0
    dd = 0
    If (x - a) * (x - b) > 0 Then
        'error:x not in range in routine chebev
        Exit Function
    End If
    y = (2 * x - a - b) / (b - a)
    y2 = 2 * y
    For j = m - 1 To 1 Step -1
        sv = d
        d = y2 * d - dd + c(j)
        dd = sv
    Next j
    getChebev = y * d - dd + 0.5 * c(0)
End Function

Private Sub getBesselIK(ByVal x As Double, ByVal xnu As Double, ByRef ri As Double, _
            ByRef rk As Double, ByRef rip As Double, ByRef rkp As Double)
    Const EPS As Double = 0.0000000001
    Const FPMIN As Double = 1E-30
    Const MAXIT As Integer = 10000
    Const XMIN As Integer = 2
    
    Dim i As Integer, l As Integer, nl As Integer
    Dim a As Double, a1 As Double, b As Double, c As Double, d As Double, del As Double, del1 As Double, delh As Double, dels As Double, e As Double, f As Double, fact As Double, fact2 As Double, ff As Double, gam1 As Double, gam2 As Double
    Dim gammi As Double, gampl As Double, h As Double, p As Double, pimu As Double, q As Double, q1 As Double, q2 As Double, qnew As Double, ril As Double, ril1 As Double, rimu As Double, rip1 As Double, ripl As Double
    Dim ritemp As Double, rk1 As Double, rkmu As Double, rkmup As Double, rktemp As Double, s As Double, sum As Double, sum1 As Double, x2 As Double, xi As Double, xi2 As Double, xmu As Double, xmu2
    
    If x <= 0 Or xnu < 0 Then
        'error:"bad arguments in bessik"
        Exit Sub
    End If
    nl = Int(xnu + 0.5)
    xmu = xnu - nl
    xmu2 = xmu * xmu
    xi = 1 / x
    xi2 = 2 * xi
    h = xnu * xi
    If h < FPMIN Then h = FPMIN
    b = xi2 * xnu
    d = 0
    c = h
    
    For i = 1 To MAXIT
        b = b + xi2
        d = 1 / (b + d)
        c = b + 1 / c
        del = c * d
        h = del * h
        If Abs(del - 1) < EPS Then Exit For
    Next i
    If (i > MAXIT) Then
        'error:("x too large in bessik; try asymptotic expansion")
        lastErr = ERR_BesselErr
        lastErrNum = ERR_BesselErrN
        Exit Sub
    End If
    ril = FPMIN
    ripl = h * ril
    ril1 = ril
    rip1 = ripl
    fact = xnu * xi

    For l = nl To 1 Step -1
        ritemp = fact * ril + ripl
        fact = fact - xi
        ripl = fact * ritemp + ril
        ril = ritemp
    Next l
    f = ripl / ril
    If (x < XMIN) Then
        'MsgBox "x<xmin =" & x
        x2 = 0.5 * x
        pimu = PI * xmu
'        fact = (fabs(pimu) < EPS ? 1.0 : pimu/sin(pimu));
        fact = IIf(Abs(pimu) < EPS, 1, pimu / Sin(pimu))
        d = -Log(x2)
        e = xmu * d
        
        'fact2 = (fabs(e) < EPS ? 1.0 : sinh(e)/e);
        fact2 = IIf(Abs(e) < EPS, 1, (Exp(e) - Exp(-e)) / (2 * e))
        
        getBeschb xmu, gam1, gam2, gampl, gammi
        ff = fact * (gam1 * (Exp(e) + Exp(-e)) / 2 + gam2 * fact2 * d)
        sum = ff
        e = Exp(e)
        p = 0.5 * e / gampl
        q = 0.5 / (e * gammi)
        c = 1
        d = x2 * x2
        sum1 = p

        For i = 1 To MAXIT
            ff = (i * ff + p + q) / (i * i - xmu2)
            c = c * (d / i)
            p = p / (i - xmu)
            q = q / (i + xmu)
            del = c * ff
            sum = sum + del
            del1 = c * (p - i * ff)
            sum1 = sum1 + del1
            If Abs(del) < Abs(sum) * EPS Then Exit For
        Next i
        If (i > MAXIT) Then
            'error:("bessk series failed to converge");
            lastErr = ERR_BesselErr
            lastErrNum = ERR_BesselErrN
            Exit Sub
        End If
        rkmu = sum
        rk1 = sum1 * xi2
    Else
        b = 2 * (1 + x)
        d = 1 / b
        h = d
        delh = d
        q1 = 0 '/*Initializations for recurrence (6.7.35).*/
        q2 = 1
        a1 = 0.25 - xmu2
        q = a1
        c = a1 '/*First term in equation (6.7.34).*/
        a = -a1
        s = 1 + q * delh
        For i = 2 To MAXIT
            a = a - 2 * (i - 1)
            c = -a * c / i
            qnew = (q1 - b * q2) / a
            q1 = q2
            q2 = qnew
            q = q + c * qnew
            b = b + 2
            d = 1 / (b + a * d)
            delh = (b * d - 1) * delh
            h = h + delh
            dels = q * delh
            s = s + dels
            If Abs(dels / s) < EPS Then Exit For
        Next i
        If (i > MAXIT) Then
            'error:("bessik: failure to converge in cf2");
            lastErr = ERR_BesselErr
            lastErrNum = ERR_BesselErrN
            Exit Sub
        End If
        h = a1 * h
        rkmu = Sqr(PI / (2 * x)) * Exp(-x) / s
        rk1 = rkmu * (xmu + x + 0.5 - h) * xi
    End If
    rkmup = xmu * xi * rkmu - rk1
    rimu = xi / (f * rkmu - rkmup)
    ri = (rimu * ril1) / ril
    rip = (rimu * rip1) / ril
    
    For i = 1 To nl
        rktemp = (xmu + i) * xi2 * rk1 + rkmu
        rkmu = rk1
        rk1 = rktemp
    Next i
    rk = rkmu
    rkp = xnu * xi * rkmu - rk1
End Sub


Public Function functionAiry(ByVal x As Double, Optional k As Integer = 0) As Double
    If k < 0 Or k > 3 Then
        lastErr = ERR_AiryErr & "The second argument must be 0, 1, 2 or 3"
        lastErrNum = ERR_AiryErrN
        Exit Function
    End If
    
    Const THIRD As Double = 1 / 3
    Const TWOTHR As Double = 2 / 3
    Const ONOVRT As Double = 0.57735027
    
    Dim ai As Double, bi As Double
    Dim aip As Double, bip As Double
    Dim absx As Double, ri As Double, rip As Double, rj As Double
    Dim rjp As Double, rk As Double, rkp As Double
    Dim rootx As Double, ry As Double, ryp As Double, z As Double

    absx = Abs(x)
    rootx = Sqr(absx)
    z = TWOTHR * absx * rootx
    If x > 0 Then
        getBesselIK z, THIRD, ri, rk, rip, rkp
        ai = rootx * ONOVRT * rk / PI
        bi = rootx * (rk / PI + 2 * ONOVRT * ri)
        getBesselIK z, TWOTHR, ri, rk, rip, rkp
        aip = -x * ONOVRT * rk / PI
        bip = x * (rk / PI + 2 * ONOVRT * ri)
    ElseIf x < 0 Then
        getBesselJY z, THIRD, rj, ry, rjp, ryp
        ai = 0.5 * rootx * (rj - ONOVRT * ry)
        bi = -0.5 * rootx * (ry + ONOVRT * rj)
        getBesselJY z, TWOTHR, rj, ry, rjp, ryp
        aip = 0.5 * absx * (ONOVRT * ry + rj)
        bip = 0.5 * absx * (ONOVRT * rj - ry)
    Else
        ai = 0.35502805
        bi = ai / ONOVRT
        aip = -0.2588194
        bip = -aip / ONOVRT
    End If
    
    If k = 0 Then
        functionAiry = ai
    ElseIf k = 1 Then
        functionAiry = aip
    ElseIf k = 2 Then
        functionAiry = bi
    ElseIf k = 3 Then
        functionAiry = bip
    End If
End Function

'functions jn(x), yn(x), and their derivatives
'   k:  0 - ret jn
'       1 - ret jn'
'       2 - ret yn
'       3 - ret yn'
Public Function functionBesselSpherical(ByVal x As Double, ByVal n As Integer, Optional k As Integer = 0) As Double
    Const RTPIO2 As Double = 1.2533141
    
    Dim factor As Double, order As Double, rj As Double, rjp As Double, ry As Double, ryp As Double
    Dim sj As Double, sy As Double, sjp As Double, syp As Double
    
    If x <= 0 Then
        'error:"bad arguments"
        lastErr = ERR_SphericalErr & "The first argument must be positive"
        lastErrNum = ERR_SphericalErrN
        Exit Function
    ElseIf k < 0 Or k > 3 Then
        lastErr = ERR_SphericalErr & "The last argument must be 0, 1, 2 or 3"
        lastErrNum = ERR_SphericalErrN
        Exit Function
    End If
    order = n + 0.5
    getBesselJY x, order, rj, ry, rjp, ryp
    factor = RTPIO2 / Sqr(x)
    sj = factor * rj
    sy = factor * ry
    sjp = factor * rjp - sj / (2 * x)
    syp = factor * ryp - sy / (2# * x)

    If k = 0 Then
        functionBesselSpherical = sj
    ElseIf k = 1 Then
        functionBesselSpherical = sjp
    ElseIf k = 2 Then
        functionBesselSpherical = sy
    ElseIf k = 3 Then
        functionBesselSpherical = syp
    End If
End Function

'Riccati-Bessel functions, Sn and Cn
'   k=0: Sn else Cn
Public Function functionBesselRiccati(ByVal x As Double, ByVal n As Integer, Optional k As Integer = 0) As Double
    If k < 0 Or k > 1 Then
        lastErr = ERR_BesselErr & "The last argument must be 0 or 1"
        lastErrNum = ERR_BesselErrN
        Exit Function
    End If
    If k = 0 Then
        functionBesselRiccati = x * functionBesselSpherical(x, n, 0)
    ElseIf k = 1 Then
        functionBesselRiccati = -x * functionBesselSpherical(x, n, 2)
    End If
End Function

