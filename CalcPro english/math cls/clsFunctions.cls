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
Private libGrp As New Collection

'za kod
Private libFCodes() As New Collection

'Private sysFunctions As New Collection
'
Public Function getFuncResult(ByVal functionName As String, ByRef fArgs() As String) As String
    Dim argCnt As Integer, tmpFName As String
    argCnt = UBound(fArgs) + 1
    tmpFName = functionName & Replace$("(N)", "N", argCnt) '(LNG_OPENST & argCnt & LNG_CLOSEST)
    
    'ako je sistemska funkcija
    '
    'ako je functionName = CPL_IMAG
    If functionName = CPL_IMAG Then
        getFuncResult = CPL_IMAG
        Exit Function
    End If
    If isInCollection(tmpFName, sysFunctions) Then
        'definiramo varijable
        Dim mCplx1 As Complex, mCplx2 As Complex
        'tražimo koja je to f-ja
        Select Case StrConv(functionName, vbLowerCase)
            Case "abs"
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im = 0 Then
                    getFuncResult = Str(Abs(Val(fArgs(0))))
                Else
                    'var1 = Sqr(mCplx1.Im ^ 2 + mCplx1.Re ^ 2)
                    getFuncResult = Str(Sqr(mCplx1.Im ^ 2 + mCplx1.Re ^ 2))
                End If
            Case "atan"
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then Err.Raise 4, "clsCalculator", "Invalid atan argument data type!"
                getFuncResult = Str(Atn(Val(fArgs(0))))
            Case "cos"
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then Err.Raise 4, "clsCalculator", "Invalid cos argument data type!"
                getFuncResult = Str(Cos(Val(fArgs(0))))
            Case "exp"
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    mCplx1 = CExp(mCplx1)
                    getFuncResult = getStrCplx(mCplx1)
                Else
                    getFuncResult = Str(Exp(Val(fArgs(0))))
                End If
            Case "fix"
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    mCplx1.Re = Fix(mCplx1.Re)
                    mCplx1.Im = Fix(mCplx1.Im)
                    getFuncResult = getStrCplx(mCplx1)
                Else
                    getFuncResult = Str(Fix(Val(fArgs(0))))
                End If
            Case "int"
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    'MsgBox "IDE"
                    mCplx1.Re = Int(mCplx1.Re)
                    mCplx1.Im = Int(mCplx1.Im)
                    getFuncResult = getStrCplx(mCplx1)
                Else
                    getFuncResult = Str(Int(Val(fArgs(0))))
                End If
            Case "log"
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    mCplx1 = CLog(mCplx1, 10)
                    getFuncResult = getStrCplx(mCplx1)
                Else
                    getFuncResult = Str(Log(Val(fArgs(0))) / Log(10))
                End If
            Case "logn"
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    mCplx1 = CLog(mCplx1, fArgs(1))
                    getFuncResult = getStrCplx(mCplx1)
                Else
                    getFuncResult = Str(Log(Val(fArgs(0))) / Log(Val(fArgs(1))))
                End If
            Case "sgn"
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then Err.Raise 4, "clsCalculator", "Invalid sgn argument data type!"
                getFuncResult = Str(Sgn(Val(fArgs(0))))
            Case "sin"
                mCplx1 = getComplex(fArgs(0))
                'MsgBox "SIN" & vbCrLf & fArgs(0)
                If mCplx1.Im <> 0 Then Err.Raise 4, "clsCalculator", "Invalid sin argument data type!"
                getFuncResult = Str(Sin(Val(fArgs(0))))
            Case "sqr"
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    mCplx1 = CSqr(mCplx1)
                    getFuncResult = getStrCplx(mCplx1)
                Else
                    If Val(fArgs(0)) >= 0 Then
                        getFuncResult = Str(Sqr(Val(fArgs(0))))
                    Else
                        mCplx1.Re = Abs(Val(fArgs(0)))
                        mCplx1.Re = Sqr(mCplx1.Re)
                        getFuncResult = Str(mCplx1.Re) & CPL_IMAG
                    End If
                End If
            Case "tan"
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then Err.Raise 4, "clsCalculator", "Invalid tan argument data type!"
                getFuncResult = Str(Tan(Val(fArgs(0))))
            Case "mod"
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then Err.Raise 4, "clsCalculator", "Invalid mod argument data type!"
                getFuncResult = Str(Val(fArgs(0)) Mod fArgs(1))
            Case "ln"
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then
                    mCplx1 = CLn(mCplx1)
                    getFuncResult = getStrCplx(mCplx1)
                Else
                    getFuncResult = Str(Log(Val(fArgs(0))))
                End If
            Case "ncr"
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Then Err.Raise 4, "clsCalculator", "Invalid mod argument data type!"
                getFuncResult = Str(combinations(fArgs(0), fArgs(1)))
             Case "npr"
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Then Err.Raise 4, "clsCalculator", "Invalid mod argument data type!"
                getFuncResult = Str(permutations(fArgs(0), fArgs(1)))
            Case "nvr"
                mCplx1 = getComplex(fArgs(0))
                mCplx2 = getComplex(fArgs(1))
                If mCplx1.Im <> 0 Or mCplx2.Im <> 0 Then Err.Raise 4, "clsCalculator", "Invalid mod argument data type!"
                getFuncResult = Str(variations(fArgs(0), fArgs(1)))
            Case "fact"
                mCplx1 = getComplex(fArgs(0))
                If mCplx1.Im <> 0 Then Err.Raise 4, "clsCalculator", "Invalid mod argument data type!"
                getFuncResult = Str(factoriel(fArgs(0)))
            'complex
            Case "im"
                mCplx1 = getComplex(fArgs(0))
                getFuncResult = Str(mCplx1.Im)
            Case "re"
                mCplx1 = getComplex(fArgs(0))
                getFuncResult = Str(mCplx1.Re)
            Case "angle"
                mCplx1 = getComplex(fArgs(0))
                getFuncResult = Str(cArg(mCplx1))
        End Select
    'pretraživanje bibloteka
    ElseIf isInCollection(tmpFName, libCatalog) Then
        getFuncResult = getLibRes(tmpFName, fArgs)
    Else
        Err.Raise 3, "clsCalculator", "Undefined function '" & functionName & "' or wrong number or arguments!"
        'MsgBox "NEMA"
    End If
    'MsgBox "OVA:" & getFuncResult
    'Set mcals = Nothing
End Function



Private Sub Class_Initialize()
        sysFunctions.Add "abs(1)", "abs(1)"
        sysFunctions.Add "atan(1)", "atan(1)"
        sysFunctions.Add "cos(1)", "cos(1)"
        sysFunctions.Add "exp(1)", "exp(1)"
        sysFunctions.Add "fix(1)", "fix(1)"
        sysFunctions.Add "int(1)", "int(1)"
        sysFunctions.Add "log(1)", "log(1)"
        sysFunctions.Add "logN(2)", "logN(2)"
        sysFunctions.Add "sgn(1)", "sgn(1)"
        sysFunctions.Add "sin(1)", "sin(1)"
        sysFunctions.Add "sqr(1)", "sqr(1)"
        sysFunctions.Add "tan(1)", "tan(1)"
        sysFunctions.Add "mod(2)", "mod(2)"
        sysFunctions.Add "ln(1)", "ln(1)"
        sysFunctions.Add "fact(1)", "fact(1)"
        sysFunctions.Add "nCr(2)", "nCr(2)"
        sysFunctions.Add "nPr(2)", "nPr(2)"
        sysFunctions.Add "nVr(2)", "nVr(2)"
        'complex
        sysFunctions.Add "im(1)", "im(1)"
        sysFunctions.Add "re(1)", "re(1)"
        sysFunctions.Add "angle(1)", "angle(1)"
End Sub

Private Sub Class_Terminate()
    'Set libGrp = Nothing
    'Set libCatalog = Nothing
    Set sysFunctions = Nothing
    Set libCatalog = Nothing
    Set libGrp = Nothing
End Sub



'////////////////////////////////////////////////////////////////////
'               rješavanje lib funkcije
'////////////////////////////////////////////////////////////////////
Private Function getLibRes(ByVal fName As String, ByRef fArgs() As String) As String
    Dim FF As Integer, strLn As String, tmpFile As String
    Dim i As Integer, j As Integer
    Dim mFnc As String, errDesc As String
    'za varijable iz define
    Dim tmpVarNms() As String, tmpVarVals() As String
    'za math line
    Dim vNam As String, retVal As String
    'za if
    Dim mIF As cmdIF, ifLvl As Integer
    'za while
    Dim mWhileStog() As cmdWhile, tmpWhile As cmdWhile, whileLvl As Integer
    'ako je na stogu 1 izvršava se, u suprotnom ne
    Dim execStog As New Collection, lvlStog As New Collection
    'return
    Dim mStrRet As String
    'ID grupe (file-a)      start u datotece f-je
    Dim grpNum As Integer, fStart As Integer, pos1 As Integer
    'calc
    Dim mCalc As New clsCalculator
    'varijable
    Dim argVarHolder As New Collection, argVarNames As New Collection
    'kod f-je
    Dim codeHolder As New Collection
    'tmp vars za count...
    Dim codeHldCnt As Integer, argsCnt As Integer
    FF = FreeFile
    '
    '********TESTING SPEED*************
'    Dim tmr1 As Double
'    tmr1 = Timer

    'naðemo poèetak f-je u datoteci i ID grupe (pocetak:ID)
    mFnc = libCatalog.Item(fName)
    'sada izdvoimo te podatke
    pos1 = InStr(1, mFnc, ":")
    fStart = Val(Mid$(mFnc, 1, pos1 - 1))
    grpNum = Val(Mid$(mFnc, pos1 + 1))
    
    'uèitavanje linija koda funkcije
    tmpFile = libGrp.Item(grpNum)
    pos1 = 0
    whileLvl = 0
    mStrRet = "0"
   ' MsgBox "LOADING"
    Open tmpFile For Input As #FF
       'odemo odma na funkciju
       Seek #FF, fStart
        Do Until EOF(FF)
            'pos1 = pos1 + 1
            Line Input #FF, strLn
            'If pos1 >= fStart Then
                'MsgBox strLn 'Seek(FF) & vbCrLf & pos1 & vbCrLf & fStart
                If StrComp(Mid$(strLn, 1, LEN_ENDFUNCT), "end function", vbTextCompare) = 0 Then
                    Exit Do
                End If
                
                ClearString strLn
                
                If strLn <> vbNullString And Mid$(strLn, 1, 2) <> LIB_COMMENT Then
                    codeHolder.Add strLn
                End If
            'End If
        Loop
    Close #FF
    'izdvajanje parametara
    mFnc = codeHolder.Item(1)
    populateWithArgVars mFnc, argVarHolder, argVarNames
    'postavljanje vrijednosti varijabli
    argsCnt = argVarHolder.Count
    For i = 1 To argsCnt
        argVarHolder.Remove (argVarNames.Item(i))
        argVarHolder.Add fArgs(i - 1), argVarNames.Item(i)
    Next i
    '
    execStog.Add 1
    lvlStog.Add 0
    ifLvl = 0
    whileLvl = 0
    codeHldCnt = codeHolder.Count
    For i = 2 To codeHldCnt
        'ako je return
        strLn = codeHolder.Item(i)
        'zamjena argumenata sa vrijednostima
        For j = 1 To argsCnt
            'ako je argument kompleksni broj, onda ga stavimo u zagrade
            If InStrB(1, argVarHolder.Item(argVarNames.Item(j)), CPL_IMAG) > 0 Then
                replaceVarVal strLn, argVarNames.Item(j), (LNG_OPENST & argVarHolder.Item(argVarNames.Item(j)) & LNG_CLOSEST)
            Else
                replaceVarVal strLn, argVarNames.Item(j), argVarHolder.Item(argVarNames.Item(j))
            End If
        Next j
        'return
        If StrComp(Mid$(strLn, 1, Len("return")), "return", vbTextCompare) = 0 Then
            'ako je kod pod blokom koji ima true
            If execStog.Item(execStog.Count) = 1 Then
                'provjera return linije
                If checkReturn(strLn, errDesc) = False Then
                    Err.Raise -11, "clsCalculator", "Sintax error:" & vbCrLf & errDesc & strLn & vbCrLf & vbCrLf & "File:" & vbCrLf & tmpFile
                End If
                strLn = Replace$(strLn, "return", vbNullString, , , vbTextCompare)
                'maknemo ; sa kraja
                If Mid$(strLn, Len(strLn), 1) = LNG_ARGDELIMITER Then
                    strLn = Mid$(strLn, 1, Len(strLn) - 1)
                End If
                'strLn = Replace$(strLn, LNG_ARGDELIMITER, VBNullstring, , , vbTextCompare)
                'brišemo sve praznine
                If InStrB(1, strLn, " ") <> 0 Then
                    strLn = Replace$(strLn, " ", vbNullString, , , vbTextCompare)
                End If
                'tražimo rez
                mStrRet = mCalc.Calculate(strLn)
                Exit For
            End If
        'define
        ElseIf StrComp(Mid$(strLn, 1, LEN_DEFINE), "define", vbTextCompare) = 0 Then
            If execStog.Item(execStog.Count) = 1 Then
                getDefineVars strLn, tmpVarNms, tmpVarVals, tmpFile
                For j = 0 To UBound(tmpVarNms)
                    mCalc.defineVariable tmpVarNms(j), tmpVarVals(j)
                Next j
            End If
        'if
        ElseIf StrComp(Mid$(strLn, 1, Len("if")), "if", vbTextCompare) = 0 Then
            ifLvl = ifLvl + 1
            'lvlStog.Add ifLvl
            'ako se if nalazi unutar bloka koji je true
            If execStog.Item(execStog.Count) = 1 Or ifLvl = 1 Then
                getIF strLn, mIF
                If mCalc.Calculate(mIF.ifCondition) = "1" Then
                    execStog.Add 1
                Else
                    execStog.Add 0
                End If
                lvlStog.Add ifLvl
            End If
        'else if
        ElseIf StrComp(Mid$(strLn, 1, Len("elseif")), "elseif", vbTextCompare) = 0 Or StrComp(Mid$(strLn, 1, Len("else if")), "else if", vbTextCompare) = 0 Then
            'ako prvobitni if nije zadovoljen i ako je ofo else od prvobitnog
            If execStog.Item(execStog.Count) = 0 And lvlStog.Item(lvlStog.Count) = ifLvl Then
                getIF strLn, mIF
                'MsgBox strLn
                If mCalc.Calculate(mIF.ifCondition) = "1" Then
                    execStog.Remove execStog.Count
                    execStog.Add 1
                    'lvlStog.Add ifLvl
                End If
            End If
        'else
        ElseIf StrComp(Mid$(strLn, 1, Len("else")), "else", vbTextCompare) = 0 Then
            'ako prvobitni if nije zadovoljen i ako je ofo else od prvobitnog
            If execStog.Item(execStog.Count) = 0 And lvlStog.Item(lvlStog.Count) = ifLvl Then
                'getIF strLn, mIF
                'MsgBox "IDE"
                'if mCalc.Calculate(mIF.ifCondition) = "1" Then
                    execStog.Remove execStog.Count
                    execStog.Add 1
                    'lvlStog.Add ifLvl
               ' End If
            End If
        'end if
        ElseIf StrComp(Mid$(strLn, 1, Len("end if")), "end if", vbTextCompare) = 0 Or StrComp(Mid$(strLn, 1, Len("endif")), "endif", vbTextCompare) = 0 Then
            'end if without if
            If execStog.Count = 1 Then
                Err.Raise -11, "clsCalculator", "End if without if" & vbCrLf & vbCrLf & "File:" & vbCrLf & tmpFile
            Else
                'provjera da li ima poèetak if-a ukoliko se nalazimo u while
                If whileLvl > 0 Then
                    If mWhileStog(UBound(mWhileStog)).mLevel > ifLvl Then
                        Err.Raise -11, "clsCalculator", "End if without block if"
                    End If
                End If
                If ifLvl = lvlStog.Item(lvlStog.Count) Then
                    lvlStog.Remove (lvlStog.Count)
                    execStog.Remove execStog.Count
                End If
                ifLvl = ifLvl - 1
            End If
        'while
        ElseIf StrComp(Mid$(strLn, 1, Len("while")), "while", vbTextCompare) = 0 Then
            If execStog.Item(execStog.Count) = 1 Then
                ifLvl = ifLvl + 1
                'tražimo parametre while petlje
                getWhile strLn, tmpWhile, tmpFile
                'postavljamo level na if level
                tmpWhile.mLevel = ifLvl
                'postavljamo broj linije od koje se izvršava kod
                tmpWhile.startLn = i + 1
                'gledamo da li je uvjet zadovoljen
                If mCalc.Calculate(tmpWhile.whileCondition) = "1" Then
                    'ako je, na execStog stavljamo 1,
                    '   stavimo while na whileStog i
                    '   poveæamo while level
                    execStog.Add 1
                    addOnWhileStrog mWhileStog, tmpWhile, whileLvl
                    whileLvl = whileLvl + 1
                Else
                    'ako nije na exec stog ide 0
                    execStog.Add 0
                End If
            End If
        'loop
        ElseIf StrComp(Mid$(strLn, 1, Len("loop")), "loop", vbTextCompare) = 0 Then
            If execStog.Item(execStog.Count) = 1 Then
                'ako je na execStogu samo jedan podatak,
                '   znaèi da imamo neku grešku u kodu
                If execStog.Count = 1 Then
                    Err.Raise -11, "clsCalculator", "Loop without while" & vbCrLf & vbCrLf & "File:" & vbCrLf & tmpFile
                ElseIf mWhileStog(UBound(mWhileStog)).mLevel < ifLvl Then
                    Err.Raise -11, "clsCalculator", "Loop without while" & vbCrLf & vbCrLf & "File:" & vbCrLf & tmpFile
                Else
                    'opet na loop provjeravamo uvjet i ako je ispunjen,
                    '   vracamo se na poèetak while petlje
                    If mCalc.Calculate(mWhileStog(UBound(mWhileStog)).whileCondition) = "1" Then
                        i = mWhileStog(UBound(mWhileStog)).startLn - 1
                        'MsgBox "VRACA:" & mWhileStog(UBound(mWhileStog)).whileCondition & vbCrLf & "I:" & i
                    Else
                        'ako nije, sa exec stoga skidamo
                        '   zanji podatak, skidamo while sa stoga
                        execStog.Remove execStog.Count
                        removeFromWhileStog mWhileStog
                        ifLvl = ifLvl - 1
                        whileLvl = whileLvl - 1
                        'MsgBox "SKIDA NA:" & ifLvl
                    End If
                End If
            Else
                ifLvl = ifLvl - 1
                execStog.Remove execStog.Count
            End If
        'break
        ElseIf StrComp(Mid$(strLn, 1, Len("break")), "break", vbTextCompare) = 0 Then
            If execStog.Item(execStog.Count) = 1 Then
                'slièno kao loop
                execStog.Remove execStog.Count
                execStog.Add 0
                removeFromWhileStog mWhileStog
                whileLvl = whileLvl - 1
            End If
        'matematièke operacije
        ElseIf InStr(1, strLn, OP_IS) > 0 Then
            'ako je unutar bloka pod true
            If execStog.Item(execStog.Count) = 1 Then
                retMathLine strLn, vNam, retVal, tmpFile
                'provjeriti ako je argument, error
                If isInCollection(vNam, argVarHolder) = True Then
                    Err.Raise -11, "clsCalculator", "Invalid use of variable '" & vNam & "'." & vbCrLf & "Function argument value can't be changed!" & vbCrLf & vbCrLf & "File." & vbCrLf & tmpFile
                End If
                mCalc.defineVariable vNam, retVal
            End If
        Else
            Err.Raise -11, "clsCalculator", "Sintax error or unknow command:" & vbCrLf & strLn
        End If
        
        'provjera da li je sve zatvoreno (if, while...)
        If i = codeHldCnt Then
            If whileLvl > 0 Then
                Err.Raise -11, "clsCalculator", "While without loop"
            ElseIf ifLvl > 0 Then
                Err.Raise -11, "clsCalculator", "ER"
            End If
        End If
    Next i
    Set mCalc = Nothing
    Set argVarHolder = Nothing
    Set codeHolder = Nothing
    
    Set argVarNames = Nothing
    
    Set execStog = Nothing
    Set lvlStog = Nothing
    
    Erase mWhileStog
    Erase tmpVarNms
    Erase tmpVarVals
    'Set loopsStack = Nothing
    getLibRes = mStrRet
'    MsgBox Timer - tmr1
End Function
'for stog - dodavanje i brisanje sa stoga
Private Sub addOnWhileStrog(ByRef mStog() As cmdWhile, ByRef newItm As cmdWhile, ByVal nextLvl As Integer)
    ReDim Preserve mStog(nextLvl)
    mStog(nextLvl) = newItm
End Sub
Private Sub removeFromWhileStog(ByRef mStog() As cmdWhile)
    Dim mStog1() As cmdWhile
    Dim i As Integer
    If UBound(mStog) > 0 Then
        ReDim mStog1(UBound(mStog) - 1)
        For i = 0 To UBound(mStog1)
            mStog1(i) = mStog(i)
        Next i
        ReDim mStog(UBound(mStog1))
        For i = 0 To UBound(mStog1)
            mStog(i) = mStog1(i)
        Next i
    Else
        ReDim mStog(0)
    End If
    Erase mStog1
End Sub

'uèitavanje biblioteke s funkcijama
Friend Sub LoadLib(ByVal strLibFile As String)
   ' MsgBox strLibFile
    Dim FF As Integer, posStart As Integer, posEnd As Integer, i As Integer
    Dim mLvl As Integer
    Dim strLn As String
    FF = FreeFile
    mLvl = 0
    'MsgBox strLibFile
    'dodavanje grupe
    libGrp.Add strLibFile, Str(libGrp.Count + 1)
    'MsgBox strLibFile
    i = 0
    Open strLibFile For Input As #FF
        Do Until EOF(FF)
            'i = Seek(FF) '**********
            Line Input #FF, strLn
            ClearString strLn
            'provjeravamo da li je komentar //, ; ili '
            If Mid$(strLn, 1, 2) <> LIB_COMMENT Then ' Or Mid$(strLn, 1, 1) <> "'" Or Mid$(strLn, 1, 1) <> LNG_ARGDELIMITER Then
                'pocetak funkcije
                If StrComp(Mid$(strLn, 1, LEN_FUNCTION), "function", vbTextCompare) = 0 Then
                    If mLvl > 0 Then
                        Err.Raise -5, "clsFunction", "Expected end function before: " & strLn & vbCrLf & vbCrLf & "File:" & vbCrLf & strLibFile
                    Else
                        mLvl = mLvl + 1
                        strLn = getFuncName(strLn)
                        'dodavanje funkciej u katalog funkcija
                        libCatalog.Add i & ":" & libGrp.Count, strLn
                        ReDim Preserve libFCodes(i)
                        learString strLn
                        libFCodes(i).Add strLn
                    End If
                'kraj funkcije
                ElseIf StrComp(Mid$(strLn, 1, LEN_ENDFUNCT), "end function", vbTextCompare) = 0 Then
                    mLvl = mLvl - 1
                    If mLvl < 0 Then
                        Err.Raise -5, "clsFunction", "Expected function before: " & strLn & vbCrLf & vbCrLf & "File:" & vbCrLf & strLibFile
                    End If
                    i = i + 1
                'kod f-je
                Else
                    ClearString strLn
                    If strLn <> vbNullString And Mid$(strLn, 1, 2) <> LIB_COMMENT Then
                        If checkCode(strLn) = "0" Then
                            libFCodes(i).Add strLn
                            'MsgBox strLn & vbCrLf & i
                        Else
                            Err.Raise -11, "clsCalculator", "last error"
                        End If
                    End If
                End If
            End If
        Loop
    Close #FF
End Sub
