VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private sysFunctions As New Collection
Private libCatalog As New Collection
Dim mFCount As Integer              'broj uèitanih f-ja

'za kod
Private libFCodes() As New Collection

'##########################################################
'#      svojstva    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
'##########################################################

'broj funkcija
Friend Property Get FunctsCount() As Integer

    FunctsCount = sysFunctions.Count
    
End Property

'naziv f-je
Friend Property Get FunctionName(ByVal FIndex As Integer) As String
    
    If FIndex <= sysFunctions.Count Then
        FunctionName = sysFunctions.Item(FIndex)
    Else
        Err.Raise 380   'invalid property value
    End If
    
End Property

'broj biblioteka
Friend Property Get FunctionLibCount() As Integer
    
    FunctionLibCount = libCatalog.Count
    
End Property

'naziv biblioteke
Friend Property Get FunctionLibName(ByVal FIndex As Integer) As String
    
    Dim tmpStr As String
    If FIndex <= libCatalog.Count Then
        
        tmpStr = libFCodes(libCatalog.Item(FIndex)).Item(1)
        tmpStr = Replace$(tmpStr, "function ", vbNullString, , , vbTextCompare)
        FunctionLibName = tmpStr
        
    Else
        Err.Raise 380   'invalid property value
    End If
    
End Property


'##########################################################
'#      glavna f-ja   #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
'##########################################################

'funkcija koja uzima naziv f-je i parametre, te
'   traži rezultat f-je
Friend Function getFuncResult(ByVal FunctionName As String, ByRef fArgs() As String) As String
    On Error GoTo errH
    Dim argCnt As Integer, tmpFName As String
    argCnt = UBound(fArgs) + 1
    '
    FunctionName = StrConv(FunctionName, vbLowerCase)
    tmpFName = FunctionName & Replace$("(N)", "N", argCnt) '(LNG_OPENST & argCnt & LNG_CLOSEST)
    
    '
    'prvo provjeravamo da li je tražena funkcija
    '   neka sistemska funkcija
    If isInCollection(tmpFName, sysFunctions) Then
        Dim mCplx1 As Complex, mCplx2 As Complex, mCplx3 As Complex
        
        'tražimo koja je to f-ja
        Select Case FunctionName 'StrConv(FunctionName, vbLowerCase)
            '--------------------------------------------------------------
            '                   trigonometrija
            '--------------------------------------------------------------
            Case "abs"
            
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im = 0 Then
                    getFuncResult = Str$(Abs(Val(fArgs(0))))
                Else
                    'var1 = Sqr(mCplx1.Im ^ 2 + mCplx1.Re ^ 2)
                    getFuncResult = Str$(Sqr(mCplx1.Im ^ 2 + mCplx1.Re ^ 2))
                End If
                
            Case "sin"
            
                mCplx1 = getComplex(fArgs(0))
                getFuncResult = getStrCplx(CSin(mCplx1))
                
            Case "tan"
            
                mCplx1 = getComplex(fArgs(0))
                getFuncResult = getStrCplx(CTan(mCplx1))
                
            Case "atan"
            
                mCplx1 = getComplex(fArgs(0))
                getFuncResult = getStrCplx(CAtan(mCplx1))
                
            Case "cos"
            
                mCplx1 = getComplex(fArgs(0))
                getFuncResult = getStrCplx(CCos(mCplx1))
                
            '--------------------------------------------------------------
            '                   exp, log
            '--------------------------------------------------------------
            Case "exp"
            
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    mCplx1 = CExp(mCplx1)
                    getFuncResult = getStrCplx(mCplx1)
                Else
                    getFuncResult = Str$(Exp(Val(fArgs(0))))
                End If
                
            Case "log"
            
                mCplx1 = getComplex(fArgs(0))
                getFuncResult = getStrCplx(CLog(mCplx1, 10))
                
            Case "logn"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                'error: invalid base
                If mCplx2.Im <> 0 Then
                    lastErr = ERR_InvalidLogBase
                    lastErrNum = ERR_InvalidLogBaseN
                    Exit Function
                ElseIf Val(fArgs(1)) <= 0 Then
                    lastErr = ERR_InvalidLogBase
                    lastErrNum = ERR_InvalidLogBaseN
                    Exit Function
                End If
                getFuncResult = getStrCplx(CLog(mCplx1, Val(fArgs(1))))
                
            Case "ln"
            
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    mCplx1 = CLn(mCplx1)
                    getFuncResult = getStrCplx(mCplx1)
                Else
                    getFuncResult = Str$(Log(Val(fArgs(0))))
                End If
                
            '--------------------------------------------------------------
            '                   number functions
            '--------------------------------------------------------------
            Case "fix"
            
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    mCplx1.Re = Fix(mCplx1.Re)
                    mCplx1.Im = Fix(mCplx1.Im)
                    getFuncResult = getStrCplx(mCplx1)
                Else
                    getFuncResult = Str$(Fix(Val(fArgs(0))))
                End If
                
            Case "int"
            
                mCplx1 = getComplex(fArgs(0))
                
                If mCplx1.Im <> 0 Then
                    mCplx1.Re = Int(mCplx1.Re)
                    mCplx1.Im = Int(mCplx1.Im)
                    getFuncResult = getStrCplx(mCplx1)
                Else
                    getFuncResult = Str$(Int(Val(fArgs(0))))
                End If
                
            Case "frac"
            
                mCplx1 = getComplex(fArgs(0))
                getFuncResult = getStrCplx(cFrac(mCplx1))
                
            Case "sgn"
            
                mCplx1 = getComplex(fArgs(0))
                'error
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(Sgn(Val(fArgs(0))))
                
            '--------------------------------------------------------------
            '                   sqr, pow, mod
            '--------------------------------------------------------------
            Case "sqr"
            
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    mCplx1 = CSqr(mCplx1)
                    getFuncResult = getStrCplx(mCplx1)
                Else
                    If Val(fArgs(0)) >= 0 Then
                        getFuncResult = Str$(Sqr(Val(fArgs(0))))
                    Else
                        mCplx1.Re = Abs(Val(fArgs(0)))
                        mCplx1.Re = Sqr(mCplx1.Re)
                        'ako nije oblika 1i
                        If mCplx1.Re <> 1 Then
                            getFuncResult = fArgs(0) & CPL_IMAG
                        Else
                            getFuncResult = CPL_IMAG
                        End If
                    End If
                End If
                
            Case "mod"
            
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                'provjera da slucajno nije dijeljene
                '   s nulom
                If Val(fArgs(1)) <> 0 Then
                    getFuncResult = Str$(Val(fArgs(0)) Mod Val(fArgs(1)))
                Else
                    lastErr = ERR_DivByZer
                    lastErrNum = ERR_DivByZerN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                
            '--------------------------------------------------------------
            '                   kombinatorika
            '--------------------------------------------------------------
            Case "ncr"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                'error
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(combinations(Val(fArgs(0)), Val(fArgs(1))))
                
             Case "npr"
             
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                'error
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(permutations(fArgs(0), fArgs(1)))
                
            Case "nvr"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                'error
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(variations(fArgs(0), fArgs(1)))
                
            Case "fact"
            
                mCplx1 = getComplex(fArgs(0))
                'error
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(factoriel(Val(fArgs(0))))
                
            '--------------------------------------------------------------
            '                   statistika
            '--------------------------------------------------------------
            Case "binompmf"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                mCplx3 = getComplex(fArgs(2))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Or mCplx3.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(BinomPMF(Val(fArgs(0)), Val(fArgs(1)), Val(fArgs(2))))
                
            Case "binomcdf"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                mCplx3 = getComplex(fArgs(2))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Or mCplx3.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(BinomCDF(Val(fArgs(0)), Val(fArgs(1)), Val(fArgs(2))))
                
            'fisherpdf
            Case "fisherpdf"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                mCplx3 = getComplex(fArgs(2))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Or mCplx3.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(FisherPDF(Val(fArgs(0)), Val(fArgs(1)), Val(fArgs(2))))
                
            'fishercdf
            Case "fishercdf"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                mCplx3 = getComplex(fArgs(2))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Or mCplx3.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
'                If Val(fArgs(0)) > 0.9 And Val(fArgs(0)) < 1.2 Then
'                    MsgBox Val(fArgs(0)) & vbCrLf & Val(fArgs(1)) & vbCrLf & Val(fArgs(2))
'                End If
                getFuncResult = Str$(FisherCDF(Val(fArgs(0)), Val(fArgs(1)), Val(fArgs(2))))
                
            'normalpdf
            Case "normalpdf"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                mCplx3 = getComplex(fArgs(2))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Or mCplx3.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(NormalPDF(Val(fArgs(0)), Val(fArgs(1)), Val(fArgs(2))))
                
            'normalcdf
            Case "normalcdf"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                mCplx3 = getComplex(fArgs(2))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Or mCplx3.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(NormalCDF(Val(fArgs(0)), Val(fArgs(1)), Val(fArgs(2))))
                
            'ChiSquareCDF '
            Case "chisquarecdf"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(ChiSquareCDF(Val(fArgs(0)), Val(fArgs(1))))
                
            'ChiSquarepdf '
            Case "chisquarepdf"
                
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(ChiSquarePDF(Val(fArgs(0)), Val(fArgs(1))))
                
            'exponentialcdf
            Case "exponentialcdf"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(ExponentialCDF(Val(fArgs(0)), Val(fArgs(1))))
                
            'exponentialpdf
            Case "exponentialcdf"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(ExponentialPDF(Val(fArgs(0)), Val(fArgs(1))))
                
            'geometricpmf
            Case "geometricpmf"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(GeometricPMF(Val(fArgs(0)), Val(fArgs(1))))
                
            'geometriccdf
            Case "geometriccdf"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(GeometricCDF(Val(fArgs(0)), Val(fArgs(1))))
                
            '--------------------------------------------------------------
            '                   complex
            '--------------------------------------------------------------
            Case "im"
            
                mCplx1 = getComplex(fArgs(0))
                getFuncResult = Str$(mCplx1.Im)
                
            Case "re"
            
                mCplx1 = getComplex(fArgs(0))
                getFuncResult = fArgs(0)
                
            Case "angle"
            
                mCplx1 = getComplex(fArgs(0))
                getFuncResult = Str$(cArg(mCplx1))
                
            Case "invert"
            
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im = 0 And mCplx1.Re = 0 Then
                    lastErr = ERR_DivByZer
                    lastErrNum = ERR_DivByZerN
                    Exit Function
                ElseIf mCplx1.Im <> 0 Then
                    mCplx1 = CDiv(createComplex(1, 0), mCplx1)
                    getFuncResult = getStrCplx(mCplx1)
                Else
                    getFuncResult = Str$(1 / mCplx1.Re)
                End If
                
            Case "conj"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx1.Im = -mCplx1.Im
                getFuncResult = getStrCplx(mCplx1)
                
            '--------------------------------------------------------------
            '                   special functions
            '--------------------------------------------------------------
            'gamma funkcije
            Case "gamma"
            
                mCplx1 = getComplex(fArgs(0))
                'error
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(functionGamma(Val(fArgs(0))))
                
            Case "gammaln"
            
                mCplx1 = getComplex(fArgs(0))
                'mCplx2 = getComplex(fArgs(1))
                'error
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(functionGammaLN(Val(fArgs(0))))
                
            Case "gammainclwr"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(functionGammaIncLwr(Val(fArgs(0)), Val(fArgs(1))))
                
            Case "gammaincupr"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(functionGammaIncUpr(Val(fArgs(0)), Val(fArgs(1))))
                
            Case "gammaincq"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(functionGammaIncQ(Val(fArgs(0)), Val(fArgs(1))))
                
            Case "gammaincp"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(functionGammaIncQ(Val(fArgs(0)), Val(fArgs(1))))
                
            'beta
            
            Case "beta"
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(functionBeta(Val(fArgs(0)), Val(fArgs(1))))
                
            Case "betaln"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(functionBetaLn(Val(fArgs(0)), Val(fArgs(1))))
                
            Case "betainc"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                mCplx3 = getComplex(fArgs(2))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Or mCplx3.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(functionBetaInc(Val(fArgs(0)), Val(fArgs(1)), Val(fArgs(2))))
                
            'error
            Case "erf"
            
                mCplx1 = getComplex(fArgs(0))
                'error
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(functionErf(Val(fArgs(0))))
                
            Case "erfc"
            
                mCplx1 = getComplex(fArgs(0))
                'error
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(functionErfC(Val(fArgs(0))))
                
            'bessel
            Case "besselj"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
'                mCplx3 = getComplex(fArgs(2))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(functionBesselJ(Val(fArgs(0)), Val(fArgs(1))))
                
            Case "bessely"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(functionBesselY(Val(fArgs(0)), Val(fArgs(1))))
                
            Case "besseli"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(functionBesselI(Val(fArgs(0)), Val(fArgs(1))))
                
            Case "besselk"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(functionBesselK(Val(fArgs(0)), Val(fArgs(1))))
                
            Case "besselriccati"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                mCplx3 = getComplex(fArgs(2))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Or mCplx3.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(functionBesselRiccati(Val(fArgs(0)), Val(fArgs(1)), Val(fArgs(2))))
                
             Case "besselspherical"
             
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                mCplx3 = getComplex(fArgs(2))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Or mCplx3.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(functionBesselSpherical(Val(fArgs(0)), Val(fArgs(1)), Val(fArgs(2))))

            Case "airy"
            
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                'error
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Str$(functionAiry(Val(fArgs(0)), Val(fArgs(1))))
            
            '
            '--------------------------------------------------------------
            '                   pretvorba kutova
            '--------------------------------------------------------------
            Case "rad2deg"
            
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    mCplx1.Re = Fix(mCplx1.Re)
                    mCplx1.Im = Fix(mCplx1.Im)
                    getFuncResult = getStrCplx(mCplx1)
                Else
                    getFuncResult = Str$(Rad2Deg(Val(fArgs(0))))
                End If
                
            Case "deg2rad"
            
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    mCplx1.Re = Fix(mCplx1.Re)
                    mCplx1.Im = Fix(mCplx1.Im)
                    getFuncResult = getStrCplx(mCplx1)
                Else
                    getFuncResult = Str$(Deg2Rad(Val(fArgs(0))))
                End If
                
            '--------------------------------------------------------------
            '                   pretvorba sustava
            '--------------------------------------------------------------
            Case "dec2bin"
            
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                getFuncResult = Dec2Bin(mCplx1.Re)
                
            Case "dec2oct"
            
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                '
                getFuncResult = Dec2Oct(mCplx1.Re)
                
            Case "dec2hex"
            
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                '
                getFuncResult = Dec2Hex(mCplx1.Re)
                
            'oct
            Case "oct2bin"
            
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                '
                getFuncResult = Oct2Bin(mCplx1.Re)
                
            Case "oct2dec"
            
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                '
                getFuncResult = Oct2Dec(mCplx1.Re)
                '
            Case "oct2hex"
            
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                '
                getFuncResult = Oct2Hex(mCplx1.Re)
                
            'hex
            Case "hex2bin"
            
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                '
                getFuncResult = Hex2Bin(fArgs(0))
                
            Case "hex2dec"
            
                mCplx1 = getComplex(fArgs(0))
               ' mCplx2 = getComplex(fArgs(1))
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                '
                getFuncResult = Hex2Dec(fArgs(0))
                
            Case "hex2oct"
            
                mCplx1 = getComplex(fArgs(0))
               ' mCplx2 = getComplex(fArgs(1))
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                '
                getFuncResult = Hex2Oct(fArgs(0))
                
            'bin
            Case "bin2hex"
            
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                '
                getFuncResult = Bin2Hex(fArgs(0))
                
            Case "bin2dec"
            
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                '
                getFuncResult = Bin2Dec(fArgs(0))
                
            Case "bin2oct"
            
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    lastErr = ERR_FunctionErr & FunctionName
                    lastErrNum = ERR_FunctionErrN
                    getFuncResult = vbNullString
                    Exit Function
                End If
                '
                getFuncResult = Bin2Oct(fArgs(0))
                
        End Select
        
    'ukoliko nije sistemska funkcija, pretražimo
    '   f-je ucitane iz biblioteke
    ElseIf isInCollection(tmpFName, libCatalog) Then
        
        getFuncResult = getLibRes(tmpFName, fArgs)
    
    'ako je i tu nema, f-ja ne postoji i
    '   kažemo error
    Else
        
        'provjeravamo da nije došlo nešto krivo u obliku _sysVar_n_.
        If InStrB(1, FunctionName, "_sysVar_") <> 0 Then
            lastErr = ERR_InvalidFunction
            lastErrNum = ERR_InvalidFunctionN
            Exit Function
        End If
        lastErr = "Undefined function '" & FunctionName & "' or wrong number of arguments!"
        lastErrNum = ERR_UndefFuncN
        Exit Function
        
    End If
    Exit Function
    
errH:
    
    Select Case Err.Number
        
        'overflow
        Case 6
            lastErr = FunctionName & ": to big argument!"
            lastErrNum = ERR_FunctionErrN
            getFuncResult = vbNullString
            
        Case Else
            
            lastErr = ERR_FunctionErr & FunctionName
            lastErrNum = ERR_FunctionErrN
            getFuncResult = vbNullString
    
    End Select
    
    
    
    
End Function


'##########################################################
'#      class   #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
'##########################################################

Private Sub Class_Initialize()
        'trigonometrijske
        sysFunctions.Add "Rad2Deg(1)", "rad2deg(1)"
        sysFunctions.Add "Deg2Rad(1)", "deg2rad(1)"
        
        sysFunctions.Add "sin(1)", "sin(1)"
        sysFunctions.Add "cos(1)", "cos(1)"
        sysFunctions.Add "tan(1)", "tan(1)"
        sysFunctions.Add "atan(1)", "atan(1)"
        'logaritamske i exp
        sysFunctions.Add "ln(1)", "ln(1)"
        sysFunctions.Add "log(1)", "log(1)"
        sysFunctions.Add "logN(2)", "logn(2)"
        sysFunctions.Add "exp(1)", "exp(1)"
        'abs ...
        sysFunctions.Add "abs(1)", "abs(1)"
        sysFunctions.Add "sqr(1)", "sqr(1)"
        sysFunctions.Add "sgn(1)", "sgn(1)"
        sysFunctions.Add "fix(1)", "fix(1)"
        sysFunctions.Add "int(1)", "int(1)"
        sysFunctions.Add "frac(1)", "frac(1)"
        sysFunctions.Add "mod(2)", "mod(2)"
        '
        'funkcije koje izvodi clsCalculator
        sysFunctions.Add "sum(3)", "sum(3)"
        '
        sysFunctions.Add "Dec2Bin(1)", "dec2bin(1)"
        sysFunctions.Add "Dec2Hex(1)", "dec2hex(1)"
        sysFunctions.Add "Dec2Oct(1)", "dec2oct(1)"
        '
        sysFunctions.Add "Oct2Dec(1)", "oct2dec(1)"
        sysFunctions.Add "Oct2Hex(1)", "oct2hex(1)"
        sysFunctions.Add "Oct2Bin(1)", "oct2bin(1)"
        '
        sysFunctions.Add "Hex2Dec(1)", "hex2dec(1)"
        sysFunctions.Add "Hex2Oct(1)", "hex2oct(1)"
        sysFunctions.Add "Hex2Bin(1)", "hex2bin(1)"
        '
        sysFunctions.Add "Bin2Dec(1)", "bin2dec(1)"
        sysFunctions.Add "Bin2Oct(1)", "bin2oct(2)"
        sysFunctions.Add "Bin2Hex(1)", "bin2hex(1)"
        
        'kombinatorika
        sysFunctions.Add "fact(1)", "fact(1)"
        sysFunctions.Add "nCr(2)", "ncr(2)"
        sysFunctions.Add "nPr(2)", "npr(2)"
        sysFunctions.Add "nVr(2)", "nvr(2)"
        '
        'statistika
        'continous
        '
        sysFunctions.Add "ChiSquareCDF(2)", "chiSquarecdf(2)"
        sysFunctions.Add "ChiSquarePDF(2)", "chiSquarepdf(2)"
        '
        sysFunctions.Add "ExponentialCDF(2)", "exponentialcdf(2)"
        sysFunctions.Add "ExponentialPDF(2)", "exponentialpdf(2)"
        '
        sysFunctions.Add "NormalPDF(3)", "normalpdf(3)"
        sysFunctions.Add "NormalCDF(3)", "normalcdf(3)"
        '
        sysFunctions.Add "FisherPDF(3)", "fisherpdf(3)"
        sysFunctions.Add "FisherCDF(3)", "fishercdf(3)"
        '
        'discrete
        sysFunctions.Add "BinomPMF(3)", "binompmf(3)"
        sysFunctions.Add "BinomCDF(3)", "binomcdf(3)"
        '
        sysFunctions.Add "GeometricPMF(2)", "geometricpmf(2)"
        sysFunctions.Add "GeometricCDF(2)", "geometriccdf(2)"
        '
        'sysFunctions.Add "HypergeometricCDF(3)", "hypergeometriccdf(3)"
        
        'complex
        sysFunctions.Add "im(1)", "im(1)"
        sysFunctions.Add "re(1)", "re(1)"
        sysFunctions.Add "angle(1)", "angle(1)"
        sysFunctions.Add "invert(1)", "invert(1)"   '1/z
        sysFunctions.Add "conj(1)", "conj(1)"       're(z) - im(z)*i
        
        'special
        sysFunctions.Add "gamma(1)", "gamma(1)"
        sysFunctions.Add "gammaLN(1)", "gammaln(1)"
        sysFunctions.Add "gammaIncLwr(2)", "gammainclwr(2)"
        sysFunctions.Add "gammaIncUpr(2)", "gammaincupr(2)"
        sysFunctions.Add "gammaIncP(2)", "gammaincp(2)"
        sysFunctions.Add "gammaIncQ(2)", "gammaincq(2)"
        '
        sysFunctions.Add "beta(2)", "beta(2)"
        sysFunctions.Add "betaInc(3)", "betainc(3)"
        sysFunctions.Add "betaLn(2)", "betaLn(2)"
        '
        sysFunctions.Add "erf(1)", "erf(1)"
        sysFunctions.Add "erfC(1)", "erfC(1)"
        'sysFunctions.Add "errorCX(1)", "errorCX(1)"
        '
        sysFunctions.Add "besselJ(2)", "besselj(2)"
        sysFunctions.Add "besselY(2)", "bessely(2)"
        sysFunctions.Add "besselI(2)", "besseli(2)"
        sysFunctions.Add "besselK(2)", "besselk(2)"
        sysFunctions.Add "besselRiccati(3)", "besselriccati(3)"
        sysFunctions.Add "besselSpherical(3)", "besselspherical(3)"
        '
        sysFunctions.Add "airy(2)", "airy(2)"
        
        'logicke
'        sysFunctions.Add "not(1)", "not(1)"
'        sysFunctions.Add "xOR(2)", "xOR(2)"
'        sysFunctions.Add "eqv(2)", "eqv(2)"
'        sysFunctions.Add "imp(2)", "imp(2)"
End Sub

Private Sub Class_Terminate()
    Set sysFunctions = Nothing
    Set libCatalog = Nothing
    '
    Erase libFCodes()
End Sub


'##########################################################
'#      lib f-ja    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
'##########################################################

'rješavanje lib funkcije
Private Function getLibRes(ByVal fName As String, ByRef fArgs() As String) As String
    
    Dim pos1 As Integer, fStart As Integer, tmpCnt As Integer, i As Integer
    Dim mFnc As String, strLn As String, varNm As String
    'objects
    Dim mCalc As New clsCalculator
    Dim argVarNames() As String
    '
    'prena nazivu f-je naðemo podatke
    '   kao što su id kolekcije
    '   u kojoj je kod
    mFnc = libCatalog.Item(fName)
    '
    'naðemo poèetak f-je (tj. indeks
    '   colekcije u kojoj se nalazi
    '   kod f-je
    fStart = Val(mFnc)
    
    'naðemo pravi izraz f-je kako bi
    '   izdvoili argumente
    mFnc = libFCodes(fStart).Item(1)
    
    'izdvajanje argumenata
    pos1 = InStr(1, mFnc, "(")
    i = InStrRev(mFnc, ")")
    argVarNames = Split(Mid$(mFnc, pos1 + 1, i - pos1 - 1), ";")
    '
    'nakon što smo pronašli nazive
    '   argumenata, sada ih definiramo
    '   kao varijable s unesenim
    '   vrijednostima
    tmpCnt = UBound(argVarNames)
    For i = 0 To tmpCnt
        mCalc.defineVariable argVarNames(i), fArgs(i)
    Next i
    '
    'izvršavanje koda
    tmpCnt = libFCodes(fStart).Count
    'For Each vItm In libFCodes(fStart)
    For i = 2 To tmpCnt
        
        strLn = libFCodes(fStart).Item(i)
        If InStrB(1, strLn, "JP:") <> 0 Then
            
            pos1 = InStr(4, strLn, ":")
            varNm = Mid$(strLn, 4, pos1 - 4)
            If mCalc.Calculate(varNm) = "1" Then
                varNm = Mid$(strLn, pos1 + 1)
                i = Val(varNm) - 1
            End If
            
        'jump force
        ElseIf InStrB(1, strLn, "JPF:") <> 0 Then
            
            varNm = Mid$(strLn, 5)
            i = Val(varNm) - 1
            
        'define
        ElseIf InStrB(1, strLn, "def:") <> 0 Then
            
            pos1 = InStr(1, strLn, "=")
            varNm = Mid$(strLn, 5, pos1 - 5)
            mCalc.defineVariable varNm, Mid$(strLn, pos1 + 1)
            
        ElseIf InStrB(1, strLn, "ret:") <> 0 Then
            
            getLibRes = mCalc.Calculate(Mid$(strLn, 5))
            Exit For
            
        ElseIf InStrB(1, strLn, "=") <> 0 Then
            
            pos1 = InStr(1, strLn, "=")
            varNm = Mid$(strLn, 1, pos1 - 1)
            mCalc.defineVariable varNm, Mid$(strLn, pos1 + 1)
            
        ElseIf InStrB(1, strLn, "message", vbTextCompare) <> 0 Then
            
            pos1 = InStr(1, strLn, "(")
            varNm = Mid$(strLn, pos1 + 1)
            pos1 = InStrRev(varNm, ")")
            varNm = Mid$(varNm, 1, pos1 - 1)
            MsgBox Replace$(varNm, "\n", vbCrLf), , fName
            
        ElseIf InStrB(1, strLn, "setError", vbTextCompare) <> 0 Then
            
            pos1 = InStr(1, strLn, "(")
            varNm = Mid$(strLn, pos1 + 1)
            pos1 = InStrRev(varNm, ")")
            varNm = Mid$(varNm, 1, pos1 - 1)
            lastErr = Replace$(varNm, "\n", vbCrLf)
            lastErrNum = -1
            Exit Function
            
        End If
        
    Next i
    
    'i na kraju malo ocitiomo nered
    '   kako ne bi trošili memoriju
    '   koju ne koristimo
    Erase argVarNames
    Set mCalc = Nothing
End Function

'////////////////////////////////////////////////////////////////////
'               rješavanje lib funkcije
'////////////////////////////////////////////////////////////////////
'Private Function getLibRes(ByVal fName As String, ByRef fArgs() As String) As String
'    Dim strLn As String
'    Dim i As Integer, j As Integer
'    Dim mFnc As String ', errDesc As String
'    'za varijable iz define
'    Dim tmpVarNms() As String, tmpVarVals() As String
'    'za math line
'    Dim vNam As String, retVal As String
'    'za if
'    Dim mIf As cmdIF, ifLvl As Integer
'    'za while
'    Dim mWhileStog() As cmdWhile, tmpWhile As cmdWhile, whileLvl As Integer
'    'ako je na stogu 1 izvršava se, u suprotnom ne
'    Dim execStog As New Collection, lvlStog As New Collection
'    'return
'    Dim mStrRet As String
'    'ID grupe (file-a)      start u datotece f-je
'    Dim fStart As Integer, pos1 As Integer
'    'calc
'    Dim mCalc As New clsCalculator
'    'varijable
'    Dim argVarHolder As New Collection, argVarNames As New Collection
'    'tmp vars za count...
'    Dim codeHldCnt As Integer, argsCnt As Integer
'
'    'naðemo poèetak f-je u datoteci i ID grupe (pocetak:ID)
'    mFnc = libCatalog.Item(fName)
'    'sada izdvoimo te podatke
'    pos1 = Instr$(1, mFnc, ":")
'    fStart = Val(Mid$(mFnc, 1, pos1 - 1))
'    'grpNum = Val(Mid$(mFnc, pos1 + 1))
'
'    whileLvl = 0
'    mStrRet = "0"
'
'    'izdvajanje parametara
'    mFnc = libFCodes(fStart).Item(1)
'    'MsgBox "IDE:" & mFnc
'    populateWithArgVars mFnc, argVarHolder, argVarNames
'    'postavljanje vrijednosti varijabli
'    argsCnt = argVarHolder.Count
'    For i = 1 To argsCnt
'        argVarHolder.Remove (argVarNames.Item(i))
'        argVarHolder.Add fArgs(i - 1), argVarNames.Item(i)
'    Next i
'    '
'    execStog.Add 1
'    lvlStog.Add 0
'    ifLvl = 0
'    whileLvl = 0
'    codeHldCnt = libFCodes(fStart).Count
'    For i = 2 To codeHldCnt
'        'ako je return
'        strLn = libFCodes(fStart).Item(i)
'        'zamjena argumenata sa vrijednostima
'        For j = 1 To argsCnt
'            'ako je argument kompleksni broj, onda ga stavimo u zagrade
'            If InStrB(1, argVarHolder.Item(argVarNames.Item(j)), CPL_IMAG) > 0 Then
'                replaceVarVal strLn, argVarNames.Item(j), (LNG_OPENST & argVarHolder.Item(argVarNames.Item(j)) & LNG_CLOSEST)
'            Else
'                replaceVarVal strLn, argVarNames.Item(j), argVarHolder.Item(argVarNames.Item(j))
'            End If
'        Next j
'        'return
'        If InStrB(1, strLn, "return", vbTextCompare) <> 0 Then ' StrComp(Mid$(strLn, 1, Len("return")), "return", vbTextCompare) = 0 Then
'            'ako je kod pod blokom koji ima true
'            If execStog.Item(execStog.Count) = 1 Then
'                strLn = Replace$(strLn, "return", vbNullString, , , vbTextCompare)
'                'maknemo ; sa kraja
'                If Mid$(strLn, Len(strLn), 1) = LNG_ARGDELIMITER Then
'                    strLn = Mid$(strLn, 1, Len(strLn) - 1)
'                End If
'                'brišemo sve praznine
'                If InStrB(1, strLn, " ") <> 0 Then
'                    strLn = Replace$(strLn, " ", vbNullString, , , vbTextCompare)
'                End If
'                'tražimo rez
'                mStrRet = mCalc.Calculate(strLn)
'                Exit For
'            End If
'        'define
'        ElseIf InStrB(1, strLn, "define", vbTextCompare) <> 0 Then 'StrComp(Mid$(strLn, 1, LEN_DEFINE), "define", vbTextCompare) = 0 Then
'            If execStog.Item(execStog.Count) = 1 Then
'                getDefineVars strLn, tmpVarNms, tmpVarVals
'                For j = 0 To UBound(tmpVarNms)
'                    mCalc.defineVariable tmpVarNms(j), tmpVarVals(j)
'                Next j
'            End If
'        'else if
'        ElseIf InStrB(1, strLn, "elseif", vbTextCompare) <> 0 Then  'StrComp(Mid$(strLn, 1, Len("elseif")), "elseif", vbTextCompare) = 0 Or StrComp(Mid$(strLn, 1, Len("else if")), "else if", vbTextCompare) = 0 Then
'            'ako prvobitni if nije zadovoljen i ako je ofo else od prvobitnog
'            If execStog.Item(execStog.Count) = 0 And lvlStog.Item(lvlStog.Count) = ifLvl Then
'                getIF strLn, mIf
'                'MsgBox strLn
'                If mCalc.Calculate(mIf.ifCondition) = "1" Then
'                    execStog.Remove execStog.Count
'                    execStog.Add 1
'                    'lvlStog.Add ifLvl
'                End If
'            End If
'        'end if
'        ElseIf InStrB(1, strLn, "endif", vbTextCompare) <> 0 Or InStrB(1, strLn, "end if", vbTextCompare) <> 0 Then  'StrComp(Mid$(strLn, 1, Len("end if")), "end if", vbTextCompare) = 0 Or StrComp(Mid$(strLn, 1, Len("endif")), "endif", vbTextCompare) = 0 Then
'            If ifLvl = lvlStog.Item(lvlStog.Count) Then
'                lvlStog.Remove (lvlStog.Count)
'                execStog.Remove execStog.Count
'            End If
'            ifLvl = ifLvl - 1
'        'if
'        ElseIf InStrB(1, strLn, "if", vbTextCompare) <> 0 Then 'StrComp(Mid$(strLn, 1, Len("if")), "if", vbTextCompare) = 0 Then
'            ifLvl = ifLvl + 1
'            'lvlStog.Add ifLvl
'            'ako se if nalazi unutar bloka koji je true
'            If execStog.Item(execStog.Count) = 1 Or ifLvl = 1 Then
'                getIF strLn, mIf
'                If mCalc.Calculate(mIf.ifCondition) = "1" Then
'                    execStog.Add 1
'                Else
'                    execStog.Add 0
'                End If
'                lvlStog.Add ifLvl
'            End If
'        'else
'        ElseIf InStrB(1, strLn, "else", vbTextCompare) <> 0 Then ''StrComp(Mid$(strLn, 1, Len("else")), "else", vbTextCompare) = 0 Then
'            'ako prvobitni if nije zadovoljen i ako je ofo else od prvobitnog
'            If execStog.Item(execStog.Count) = 0 And lvlStog.Item(lvlStog.Count) = ifLvl Then
'                'getIF strLn, mIF
'                'MsgBox "IDE"
'                'if mCalc.Calculate(mIF.ifCondition) = "1" Then
'                    execStog.Remove execStog.Count
'                    execStog.Add 1
'                    'lvlStog.Add ifLvl
'               ' End If
'            End If
'        'while
'        ElseIf InStrB(1, strLn, "while", vbTextCompare) <> 0 Then 'StrComp(Mid$(strLn, 1, Len("while")), "while", vbTextCompare) = 0 Then
'            If execStog.Item(execStog.Count) = 1 Then
'                ifLvl = ifLvl + 1
'                'tražimo parametre while petlje
'                getWhile strLn, tmpWhile
'                'postavljamo level na if level
'                tmpWhile.mLevel = ifLvl
'                'postavljamo broj linije od koje se izvršava kod
'                tmpWhile.startLn = i + 1
'                'gledamo da li je uvjet zadovoljen
'                If mCalc.Calculate(tmpWhile.whileCondition) = "1" Then
'                    'ako je, na execStog stavljamo 1,
'                    '   stavimo while na whileStog i
'                    '   poveæamo while level
'                    execStog.Add 1
'                    addOnWhileStrog mWhileStog, tmpWhile, whileLvl
'                    whileLvl = whileLvl + 1
'                Else
'                    'ako nije na exec stog ide 0
'                    execStog.Add 0
'                End If
'            End If
'        'loop
'        ElseIf InStrB(1, strLn, "loop", vbTextCompare) <> 0 Then 'StrComp(Mid$(strLn, 1, Len("loop")), "loop", vbTextCompare) = 0 Then
'            If execStog.Item(execStog.Count) = 1 Then
'                'opet na loop provjeravamo uvjet i ako je ispunjen,
'                '   vracamo se na poèetak while petlje
'                If mCalc.Calculate(mWhileStog(UBound(mWhileStog)).whileCondition) = "1" Then
'                    i = mWhileStog(UBound(mWhileStog)).startLn - 1
'                    'MsgBox "VRACA:" & mWhileStog(UBound(mWhileStog)).whileCondition & vbCrLf & "I:" & i
'                Else
'                    'ako nije, sa exec stoga skidamo
'                    '   zanji podatak, skidamo while sa stoga
'                    execStog.Remove execStog.Count
'                    removeFromWhileStog mWhileStog
'                    ifLvl = ifLvl - 1
'                    whileLvl = whileLvl - 1
'                    'MsgBox "SKIDA NA:" & ifLvl
'                End If
'            Else
'                ifLvl = ifLvl - 1
'                execStog.Remove execStog.Count
'            End If
'        'break
'        ElseIf InStrB(1, strLn, "break", vbTextCompare) <> 0 Then ''StrComp(Mid$(strLn, 1, Len("break")), "break", vbTextCompare) = 0 Then
'            If execStog.Item(execStog.Count) = 1 Then
'                'slièno kao loop
'                execStog.Remove execStog.Count
'                execStog.Add 0
'                removeFromWhileStog mWhileStog
'                whileLvl = whileLvl - 1
'            End If
'        'matematièke operacije
'        ElseIf InStrB(1, strLn, OP_IS) > 0 Then
'            'ako je unutar bloka pod true
'            If execStog.Item(execStog.Count) = 1 Then
'                retMathLine strLn, vNam, retVal
'                'provjeriti ako je argument, error
'                If isInCollection(vNam, argVarHolder) = True Then
'                    'Err.Raise -11, "clsCalculator", "Invalid use of variable '" & vNam & "'." & vbCrLf & "Function argument value can't be changed!" & vbCrLf & vbCrLf & "File." & vbCrLf & tmpFile
'                End If
'                'MsgBox "DEF"
'                mCalc.defineVariable vNam, retVal
'            End If
''        Else
''            Err.Raise -11, "clsCalculator", "Sintax error or unknow command:" & vbCrLf & strLn
'        End If
'
'        'provjera da li je sve zatvoreno (if, while...)
''        If i = codeHldCnt Then
''            If whileLvl > 0 Then
''                Err.Raise -11, "clsCalculator", "While without loop"
''            ElseIf ifLvl > 0 Then
''                Err.Raise -11, "clsCalculator", "ER"
''            End If
''        End If
'    Next i
'    Set mCalc = Nothing
'    Set argVarHolder = Nothing
'    'Set libFCodes(fStart) = Nothing
'
'    Set argVarNames = Nothing
'
'    Set execStog = Nothing
'    Set lvlStog = Nothing
'
'    Erase mWhileStog
'    Erase tmpVarNms
'    Erase tmpVarVals
'    'Set loopsStack = Nothing
'    getLibRes = mStrRet
''    MsgBox Timer - tmr1
'End Function
'for stog - dodavanje i brisanje sa stoga
'Private Sub addOnWhileStrog(ByRef mStog() As cmdWhile, ByRef newItm As cmdWhile, ByVal nextLvl As Integer)
'    ReDim Preserve mStog(nextLvl)
'    mStog(nextLvl) = newItm
'End Sub
'Private Sub removeFromWhileStog(ByRef mStog() As cmdWhile)
'    Dim mStog1() As cmdWhile
'    Dim i As Integer
'    If UBound(mStog) > 0 Then
'        ReDim mStog1(UBound(mStog) - 1)
'        For i = 0 To UBound(mStog1)
'            mStog1(i) = mStog(i)
'        Next i
'        ReDim mStog(UBound(mStog1))
'        For i = 0 To UBound(mStog1)
'            mStog(i) = mStog1(i)
'        Next i
'    Else
'        ReDim mStog(0)
'    End If
'    Erase mStog1
'End Sub


'##########################################################
'#      ucitavanje lib f-je   #~~~~~~~~~~~~~~~~~~~~~~~~~~~#
'##########################################################

'uèitavanje biblioteke s funkcijama
Friend Sub LoadLib(ByVal strLibFile As String)
    
    Dim ff As Integer, i As Integer
    Dim mLvl As Integer
    Dim strLn As String, fNm As String
    'provjera koda (leveli)
    Dim ifLvl As Integer, whileLvl As Integer
    Dim lvlStog As New Collection
    '
    Dim tmpCode As New Collection
    ff = FreeFile
    mLvl = 0

    'otvaramo datoteku s kodom i sekvencijalno
    '   citamo liniju po liniju
    i = 0
    Open strLibFile For Input As #ff
        Do Until EOF(ff)
            
            i = i + 1
            Line Input #ff, strLn
            ClearString strLn
            strLn = Trim$(strLn)
            'provjeravamo da li je komentar // ili prazna linija
            If Mid$(strLn, 1, 2) <> LIB_COMMENT And Len(strLn) <> 0 Then ' Or Mid$(strLn, 1, 1) <> "'" Or Mid$(strLn, 1, 1) <> LNG_ARGDELIMITER Then
                
                'prvo provjerimo sintaksu, ako je nešto
                '   pogrešno izbacimo error,
                '   ovo je pretprovjera, kasnije još
                '   imamo provjeru gdje se provjerava
                '   da li su zatvorene sve petlje...
                If checkLineSintax(strLn, ifLvl, whileLvl, lvlStog) = False Then
                    
                    Err.Raise lastErrNum, "clsCalculator", lastErr & vbCrLf & vbCrLf & "File: " & right$(strLibFile, Len(strLibFile) - InStrRev(strLibFile, "\")) & vbCrLf & "Line: " & i
                
                End If
                
                
                'pocetak funkcije
                If StrComp(Mid$(strLn, 1, LEN_FUNCTION), "function", vbTextCompare) = 0 Then
                    
                    If mLvl > 0 Then
                        
                        lastErr = "Expected function before: " & strLn
                        lastErrNum = ERR_ExpectedFuncN
                        Err.Raise lastErrNum, "clsFunction", "Expected end function before: " & vbCrLf & strLn & vbCrLf & vbCrLf & "File: " & right$(strLibFile, Len(strLibFile) - InStrRev(strLibFile, "\")) & vbCrLf & "Line: " & i
                    
                    Else
'                        Do While tmpCode.Count > 0
'                            tmpCode.Remove 1
'                        Loop
                        'Set tmpCode = New Collection
                        
                        ReDim Preserve libFCodes(mFCount)
                        ClearString strLn
                        tmpCode.Add strLn
                        
                        'libFCodes(mFCount).Add strLn
                        'MsgBox mFCount & vbCrLf & strLn
                        mLvl = mLvl + 1
                        
                        'strLn = getFuncName(strLn)
                        fNm = getFuncName(strLn) 'strLn
                        'dodavanje funkciej u katalog funkcija
                        'libCatalog.Add mFCount & ":" & libGrp.Count, strLn
                        If isInCollection(StrConv(fNm, vbLowerCase), libCatalog) = False Then
                        
                            libCatalog.Add mFCount, StrConv(fNm, vbLowerCase)
                            
                        Else
                            
                            lastErr = Replace$(ERR_FunctionExist, "$fname", strLn) '"Invalid function '" & fNm & "', function already exist!"
                            lastErrNum = ERR_FunctionExistN
                            Err.Raise lastErrNum, "clsFunction", lastErr & vbCrLf & vbCrLf & "File: " & right$(strLibFile, Len(strLibFile) - InStrRev(strLibFile, "\")) & vbCrLf & "Line: " & i
                            
                        End If
                        '
                        ifLvl = 0
                        whileLvl = 0
                        
                    End If
                    
                'kraj funkcije
                ElseIf StrComp(Mid$(strLn, 1, LEN_ENDFUNCT), "end function", vbTextCompare) = 0 Then
                    
                    mLvl = mLvl - 1
                    'provjeravanje da li je sve zatvoreno
                    '   tj. da li su zatvorene sve while petlje i
                    '   if naredbe
                    If ifLvl > 0 Then
                        lastErr = ERR_ExpectedEndIf & ":" & vbCrLf & strLn
                        lastErrNum = ERR_ExpectedEndIfN
                        Err.Raise ERR_ExpectedEndIfN, "clsCalculator", ERR_ExpectedEndIf & ":" & vbCrLf & strLn & vbCrLf & vbCrLf & "File: " & right$(strLibFile, Len(strLibFile) - InStrRev(strLibFile, "\")) & vbCrLf & "Line: " & i
                    
                    ElseIf whileLvl > 0 Then
                        lastErr = ERR_ExpectedLoop & ":" & vbCrLf & strLn
                        lastErrNum = ERR_ExpectedLoopN
                        Err.Raise ERR_ExpectedLoopN, "clsCalculator", ERR_ExpectedLoop & ":" & vbCrLf & strLn & vbCrLf & vbCrLf & "File: " & right$(strLibFile, Len(strLibFile) - InStrRev(strLibFile, "\")) & vbCrLf & "Line: " & i
                    
                    ElseIf mLvl < 0 Then
                        lastErr = "Expected function before: " & strLn
                        lastErrNum = ERR_ExpectedFuncN
                        Err.Raise lastErrNum, "clsFunction", "Expected function before: " & strLn & vbCrLf & vbCrLf & "File: " & vbCrLf & strLibFile
                    
                    End If
                    'konvertiramo kod u kod prepoznatljiv
                    '   interpreteru, kod koji je slican
                    '   asembleru
                    convertToExecutable tmpCode, libFCodes(mFCount)
                    
                    'brisanje izvornog koda f-je
                    '   kako bi se u kolekciju
                    '   mogao pisati novi kod
                    Set tmpCode = Nothing
                    
                    'povecavamo broj f-ja za 1
                    mFCount = mFCount + 1
                    
                'kod f-je
                ElseIf mLvl = 0 Then
                    
                    lastErr = ERR_CodeOutside & ":" & vbCrLf & strLn
                    lastErrNum = ERR_CodeOutsideN
                    Err.Raise ERR_CodeOutsideN, "clsCalculator", ERR_CodeOutside & ":" & vbCrLf & strLn & vbCrLf & vbCrLf & "File: " & right$(strLibFile, Len(strLibFile) - InStrRev(strLibFile, "\")) & vbCrLf & "Line: " & i
                
                Else
                    
                    ClearString strLn
                    tmpCode.Add strLn
                    
                End If
                
            End If
        Loop
        
    Close #ff
    
    Set tmpCode = Nothing
    Set lvlStog = Nothing
    Exit Sub
    
'errH:
'    Set tmpCode = Nothing
'    Set lvlStog = Nothing
'    lastErr = ERR_FileError & vbCrLf & Mid$(strLibFile, InStrRev(strLibFile, "\") + 1)
'    lastErrNum = ERR_FileErrorN
'    Err.Raise lastErrNum, , lastErr
End Sub

