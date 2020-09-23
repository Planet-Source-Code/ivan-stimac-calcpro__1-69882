Attribute VB_Name = "modComplex"
Option Explicit
'//////////////////////////////////////////////////////////////////////
'                           complex
'//////////////////////////////////////////////////////////////////////
'Public Function getSqrt(ByRef z As Complex) As Complex
'    Dim k As Complex
'    z = getComplex(fArgs(0))
'    If z.Im <> 0 Then
'        getSqrt = CSqr(z)
'    Else
'        If z.Re >= 0 Then
'            getSqrt.Re = Sqr(z.Re)
'            getSqrt.Im = 0
'        Else
'            k.Re = Abs(z.Re)
'            k.Re = Sqr(k.Re)
'        End If
'    End If
'End Function

Public Function createComplex(ByVal zRe As Double, ByVal zIm As Double) As Complex

    createComplex.Re = zRe
    createComplex.Im = zIm
    
End Function

Public Function getStrCplx(ByRef z As Complex) As String

    If z.Re = 0 And z.Im = 0 Then
    
        getStrCplx = "0"
        
    ElseIf z.Re = 0 Then
    
        If z.Im <> 1 Then
            getStrCplx = z.Im & CPL_IMAG
        Else
            getStrCplx = CPL_IMAG
        End If
        
    ElseIf z.Im = 0 Then
    
        getStrCplx = z.Re
        
    Else
    
        If z.Im > 0 Then
            
            If z.Im <> 1 Then
                getStrCplx = (z.Re & OP_PLUS) & (z.Im & CPL_IMAG)
            Else
                getStrCplx = (z.Re & OP_PLUS) & CPL_IMAG
            End If
            
        Else
            
            If z.Im <> -1 Then
                getStrCplx = z.Re & (z.Im & CPL_IMAG)
            Else
                getStrCplx = z.Re & ("-" & CPL_IMAG)
            End If
            
        End If
        
    End If
End Function

Public Function getComplex(ByVal strVal As String) As Complex
    
    'ako nije kompleksni broj, igmah
    '   damo vrijednost
    If InStrB(1, strVal, CPL_IMAG) = 0 Then
    
        getComplex.Re = Val(strVal)
        getComplex.Im = 0
        Exit Function
        
    End If
    
    Dim pos As Integer, prev As Integer, nxt As Integer
    Dim mRe As String, mIm As String, strOp As String
    pos = InStr(1, strVal, CPL_IMAG, vbTextCompare)
    '
    If pos > 0 Then
        
        prev = InStrRev(strVal, OP_MINUS, pos)
        If prev < InStrRev(strVal, OP_PLUS, pos) Then
            prev = InStrRev(strVal, OP_PLUS, pos)
        End If
        
    End If
    
    nxt = getNextOperator(strVal, pos)
    If nxt = 0 Then
    
        nxt = Len(strVal)
        
    End If
    If pos > 0 Then
        
        mIm = "1"
        strOp = vbNullString
        
        If prev < 2 Then
            
            If prev = 1 Then
                strOp = Mid$(strVal, 1, 1)
            End If
            If strOp = OP_PLUS Then
                mIm = Mid$(strVal, 2, pos - 1)
            Else
                mIm = Mid$(strVal, 1, pos - 1)
            End If
            If pos - prev <= 1 Then
                mIm = strOp & "1"
            End If
            strVal = Mid$(strVal, pos + 1)
            
        Else
        
            If prev > 0 Then
                strOp = Mid$(strVal, prev, 1)
            End If
            If strOp = OP_MINUS Then
                mIm = OP_MINUS & Mid$(strVal, prev + 1, pos - prev - 1)
            Else
                mIm = Mid$(strVal, prev + 1, pos - prev - 1)
            End If
            
        End If
        
        
        If pos - prev <= 1 Then
            mIm = strOp & "1"
        End If
        
        If LenB(strVal) <> 0 And prev > 1 Then
            strVal = Mid$(strVal, 1, prev - 1)
        End If
        
    End If
    
    mRe = strVal
    If Mid$(mRe, 1, 1) = OP_PLUS Then
    
        mRe = Mid$(mRe, 2)
        
    End If
    
    
    If Mid$(mIm, 1, 1) = OP_PLUS Then
    
        mIm = Mid$(mIm, 2)
        
    End If
    
    getComplex.Re = Val(mRe)
    getComplex.Im = Val(mIm)
    
End Function

'add
Public Function CAdd(ByRef a As Complex, ByRef b As Complex) As Complex

    CAdd.Re = a.Re + b.Re
    CAdd.Im = a.Im + b.Im
    
End Function
'sub
Public Function CSub(ByRef a As Complex, ByRef b As Complex) As Complex

    CSub.Re = a.Re - b.Re
    CSub.Im = a.Im - b.Im
    
End Function
'mul
Public Function CMul(ByRef a As Complex, ByRef b As Complex) As Complex

    CMul.Re = a.Re * b.Re - a.Im * b.Im
    CMul.Im = a.Im * b.Re + a.Re * b.Im
    
End Function
'div
Public Function CDiv(ByRef a As Complex, ByRef b As Complex) As Complex

    Dim c As Complex
    Dim r As Double, den As Double
    
    If Abs(b.Re) >= Abs(b.Im) Then
    
        r = b.Im / b.Re
        den = b.Re + r * b.Im
        c.Re = (a.Re + r * a.Im) / den
        c.Im = (a.Im - r * a.Re) / den
        
    Else
    
        r = b.Re / b.Im
        den = b.Im + r * b.Re
        c.Re = (a.Re * r + a.Im) / den
        c.Im = (a.Im * r - a.Re) / den
        
    End If
    
    CDiv = c
End Function
'
Public Function CSqr(ByRef z As Complex) As Complex

    Dim c As Complex
    Dim x As Double, y As Double, w As Double, r As Double
    
    If ((z.Re = 0) And (z.Im = 0)) Then
    
        c.Re = 0
        c.Im = 0
        
    Else
    
        x = Abs(z.Re)
        y = Abs(z.Im)
        
        If (x >= y) Then
        
            r = y / x
            w = Sqr(x) * Sqr(0.5 * (1 + Sqr(1 + r * r)))
            
        Else
        
            r = x / y
            w = Sqr(y) * Sqr(0.5 * (r + Sqr(1 + r * r)))
            
        End If
        
        
        If (z.Re >= 0) Then
        
            c.Re = w
            c.Im = z.Im / (2 * w)
            
        Else
        
            c.Im = IIf(z.Im >= 0, w, -w)
            c.Re = z.Im / (2 * c.Im)
            
        End If
        
    End If
    
    CSqr = c
End Function
'potention z^n
Public Function CPot(ByRef z As Complex, ByVal n As Double) As Complex

    Dim r As Double, fi As Double
    
    r = Sqr(z.Im ^ 2 + z.Re ^ 2)
    fi = cArg(z) 'Atn(z.Im / z.Re)
    CPot.Re = r ^ n * Cos(n * fi)
    CPot.Im = r ^ n * Sin(n * fi)
    
End Function

'potention z1^z2
Public Function CPot2(ByRef z As Complex, ByRef z2 As Complex) As Complex

    Dim r As Double, fi As Double
    Dim tmpZ As Complex
    r = Sqr(z.Im ^ 2 + z.Re ^ 2)
    fi = Atn(z.Im / z.Re)
    
    
    tmpZ.Re = Log(r) '/ Log(10)
    tmpZ.Im = fi
    tmpZ = CMul(tmpZ, z2)
    CPot2 = CExp(tmpZ)
    
End Function


'
'ln(z)
'Public Function CLn(ByRef z As Complex) As Complex
'    Dim r As Double, fi As Double
'
'    r = Sqr(z.Im ^ 2 + z.Re ^ 2)
'    fi = Atn(z.Im / z.Re)
'
'    CLog.Re = Log(r)
'    CLog.Im = fi
'End Function
'




Public Function cAbs(ByRef z As Complex) As Double

    cAbs = Sqr(z.Im ^ 2 + z.Re ^ 2)
    
End Function

Public Function cArg(ByRef z As Complex) As Double

    If z.Im = 0 Then
    
        If z.Re >= 0 Then
            cArg = 0
        ElseIf z.Re < 0 Then
            cArg = PI
        End If
        
    ElseIf z.Re = 0 Then
    
        If z.Im = 0 Then
            cArg = 0
        ElseIf z.Im > 0 Then
            cArg = PI / 2
        Else
            cArg = -PI / 2
        End If
        
    Else
    
        cArg = Atn(z.Im / z.Re)
        
    End If
    
End Function

Public Function CInv(ByRef z As Complex) As Complex

    CInv = CDiv(createComplex(1, 0), z)
    
End Function


Public Function CConj(ByRef z As Complex) As Complex

    CConj.Re = z.Re
    CConj.Im = -z.Im
    
End Function
