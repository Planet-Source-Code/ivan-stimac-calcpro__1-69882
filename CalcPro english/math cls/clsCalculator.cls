VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Private mMode As eMode
Private sysVars As New Collection
Private userVars As New Collection, userVarNames As New Collection
Private mFunc As New clsFunctions
'
Private operators() As String
'consts
Private Const opBound   As Integer = 12                     'operators count
Private Const OP_LIST   As String = "+-*/%^<=>=!==&&||Ee"   'za provjeru znaka ispred operatora
Private Const OP_CANBEPREV   As String = "+-*/%^<=>=!==&&||Ee("
'Private Const CHR_SOLVEFUC As String = "_$\-+."
'Private Const CHR_SOLVEFUC2 As String = "_$\-+.i"

'##########################################################
'#      svojstva    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
'##########################################################
'popis sys funkcija
Public Property Get FunctsCount() As Integer
    
    FunctsCount = mFunc.FunctsCount
    
End Property

'naziv funkcije
Public Property Get FunctionName(ByVal FIndex As Integer) As String
    
    Dim pos1 As Integer, pos2 As Integer, num1 As Integer
    FunctionName = mFunc.FunctionName(FIndex)
    '
    pos1 = InStrRev(FunctionName, "(")
    pos2 = InStr(pos1, FunctionName, ")")
    num1 = Val(Mid$(FunctionName, pos1 + 1, pos2 - pos1 - 1))
    '
    FunctionName = Mid$(FunctionName, 1, pos1 - 1) & " ("
    '
    For pos1 = 1 To num1
        'ako je prvi parametar, bit ce x
        If pos1 > 1 Then
            FunctionName = FunctionName & ("$" & Chr(95 + pos1))
        Else
            FunctionName = FunctionName & "$x"
        End If
        If pos1 < num1 Then
            FunctionName = FunctionName & " ; "
        End If
    Next pos1
    FunctionName = FunctionName & ")"
    
End Property

'popis lib funkcija
Public Property Get FunctsLibCount() As Integer

    FunctsLibCount = mFunc.FunctionLibCount
    
End Property

Public Property Get FunctionLibName(ByVal FIndex As Integer) As String

    FunctionLibName = mFunc.FunctionLibName(FIndex)
    
End Property


'svojstva varijable(i)
Public Property Get VarCount() As Integer

    VarCount = userVarNames.Count
    
End Property

Public Property Get varName(ByVal varInd As Integer) As String

    varName = userVarNames.Item(varInd)
    
End Property

Public Property Get VarValue(ByVal varName As String) As String
                
    VarValue = userVars.Item(varName)
    
End Property
'
'##########################################################
'#      public sub    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
'##########################################################
Public Sub LoadVarLib(ByVal libFile As String)
    
    On Error GoTo errH
    Dim ff As Integer, strLn As String
    Dim pos1 As Integer
    Dim varName As String, varVal As String
    ff = FreeFile
    '
    Open libFile For Input As #ff
        
        Do Until EOF(ff)
            
            Line Input #ff, strLn
            ClearString strLn
            
            'ako je sigurno varijabla
            If left$(strLn, 1) = "$" Then
                
                pos1 = InStr(1, strLn, "=")
                If pos1 > 1 Then
                    varName = Trim$(left$(strLn, pos1 - 1))
                    varVal = Trim$(Mid$(strLn, pos1 + 1))
                    Me.defineVariable varName, varVal
                End If
                
            End If
            
        Loop
        
    Close #ff
    
    Exit Sub
    
errH:
    lastErr = ERR_FileError
    lastErrNum = ERR_FileErrorN
    Err.Raise lastErrNum, , lastErr
End Sub

'ukljucivanje biblioteke
Public Sub IncludeLib(ByVal strLib As String)

    mFunc.LoadLib strLib

End Sub

'brisanje varijabli
Public Sub ClearParseTree()

    Set sysVars = Nothing
    
End Sub

'definiranje varijable
Public Sub defineVariable(ByVal varName As String, ByVal VarValue As String)
    
    varName = Trim$(varName)
    VarValue = Trim$(VarValue)
    'prvo provjerimo ispravnost naziva
    '   varijable
    If InStrB(1, varName, ".") <> 0 Then
        lastErr = ERR_InvalidVarName
        lastErrNum = ERR_InvalidVarNameN
        Exit Sub
        
    ElseIf InStrB(1, varName, ";") <> 0 Then
        lastErr = ERR_InvalidVarName
        lastErrNum = ERR_InvalidVarNameN
        Exit Sub
        
    ElseIf InStrB(1, varName, "\") <> 0 Then
        lastErr = ERR_InvalidVarName
        lastErrNum = ERR_InvalidVarNameN
        Exit Sub
        
    ElseIf InStrB(1, varName, "+") <> 0 Or InStrB(1, varName, "-") <> 0 Or InStrB(1, varName, "/") <> 0 Or InStrB(1, varName, "*") <> 0 Or InStrB(1, varName, "%") <> 0 Or InStrB(1, varName, "^") <> 0 Or InStrB(1, varName, "&") <> 0 Or InStrB(1, varName, "=") <> 0 Or InStrB(1, varName, "|") <> 0 Or InStrB(1, varName, "!") <> 0 Then
        lastErr = ERR_InvalidVarName
        lastErrNum = ERR_InvalidVarNameN
        Exit Sub
        
    End If
    
    
    'ako je vrijednos varijable izraz,
    '   tada racunamo vrijednost
    If IsNumeric(VarValue) <> True Then
    
        VarValue = Calculate(VarValue)
        
    End If
    
    
    'ako korisnik nije stavio ispred $,
    '   mi cemo to napraviti
    If left$(varName, 1) <> LNG_DOLLAR Then
    
        varName = LNG_DOLLAR & varName
        
    End If
    
    
    'ako ta varijabla vec postoji,
    '   tada je brišemo kako bi dodali
    '   novu vrijednost
    If userVars.Count > 0 Then
    
        If checkVar(varName) = True Then
        
            userVars.Remove (varName)
            userVarNames.Remove (varName)
            
        End If
        
    End If
    
    userVars.Add VarValue, varName
    userVarNames.Add varName, varName

End Sub

'
'##########################################################
'#      public function    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
'##########################################################

'glavna funkcija za poziv rješavanja
'   funkcije
Public Function Calculate(ByVal strFunction As String) As String
    
    Dim tmpStr As String, prevChr As String
    Dim mPos As Integer, mPos1 As Integer
    
    'provjera sintakse unesene funkcije
    If checkFunction(strFunction) <> 0 Then
        
        Select Case checkFunction(strFunction)
            
            Case 1
                lastErr = ERR_ExpEndOfSt
                lastErrNum = ERR_ExpEndOfStN
                Exit Function
                
            Case 2
                lastErr = ERR_InvalidChar
                lastErrNum = ERR_InvalidCharN
                Exit Function
                
        End Select
        
    End If
    '
    'brišemo sve praznine
    ClearString strFunction
    
    If InStrB(1, strFunction, " ") <> 0 Then
    
        strFunction = Replace$(strFunction, " ", vbNullString)
        
    End If
    
    'ako nema ništa za racunati, vratimo 0
    If Len(strFunction) = 0 Then
    
        Calculate = "0"
        Exit Function
        
    End If
    '
    mPos = InStr(1, strFunction, LNG_CLOSEST)
    
    Do While mPos > 0   'dok ima zatvorenih zagrada
    
        mPos1 = InStrRev(strFunction, LNG_OPENST, mPos)
        tmpStr = Mid$(strFunction, mPos1 + 1, mPos - mPos1 - 1)
        sysVars.Add getSubRes(tmpStr), LNG_SYSVAR & (sysVars.Count + 1 & LNG_DOWNLN)
        
        If mPos1 > 0 Then
            
            'provjera da li se radi o parametru funkcije
            '   ili ne, ako je parametar, tada se stavlja
            '   u zagrade \param\
            If mPos1 <= 1 Then
                
                strFunction = Mid$(strFunction, 1, mPos1 - 1) & LNG_SYSVAR & sysVars.Count & ("_" & Mid$(strFunction, mPos + 1))
            
            Else
                
                prevChr = Mid$(strFunction, mPos1 - 1, 1)
                If InStrB(1, OP_CANBEPREV, prevChr) = 0 Then
                    strFunction = (Mid$(strFunction, 1, mPos1 - 1) & LNG_REPLACEB) & LNG_SYSVAR & sysVars.Count & ("_\" & Mid$(strFunction, mPos + 1))
                Else
                    strFunction = Mid$(strFunction, 1, mPos1 - 1) & LNG_SYSVAR & sysVars.Count & ("_" & Mid$(strFunction, mPos + 1))
                End If
                
            End If
            
        Else
        
            strFunction = sysVars.Item(LNG_SYSVAR & sysVars.Count & LNG_DOWNLN)
            Exit Do
            
        End If
        mPos = InStr(1, strFunction, LNG_CLOSEST)
        
    Loop
    Calculate = getSubRes(strFunction)

    'brisanje praznina
    If InStrB(1, Calculate, " ") <> 0 Then
        
        Calculate = Replace$(Calculate, " ", vbNullString)
        
    End If
    
    'zamjena , sa .
    If InStrB(1, Calculate, ",") <> 0 Then
        
        Calculate = Replace$(Calculate, ",", ".")
        
    End If
    
    'ako je rezultat oblika .545654, dodamo 0 na poèetak
    If Mid$(Calculate, 1, 1) = "." Then
        
        Calculate = "0" & Calculate
        
    ElseIf Mid$(Calculate, 2, 1) = "." And Mid$(Calculate, 1, 1) = OP_MINUS Then
        
        Calculate = "-0" & Mid$(Calculate, 2)
        
    End If
    
    
    '
'    Do While sysVars.Count > 0
'        sysVars.Remove 1
'    Loop
End Function

'errors
Public Function ErrorGetNumber() As Long

    ErrorGetNumber = lastErrNum
    
End Function

Public Function ErrorGetDescription() As String

    ErrorGetDescription = lastErr
    
End Function

Public Function ErrorClear() As String

    lastErr = vbNullString
    lastErrNum = 0
    
End Function

'##########################################################
'#      sub res    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
'##########################################################

'funkcija vraca rezultat izraza koji je na istom nivou:
'   izraz unutar iste zagrade (npr 2+5*8)
Private Function getSubRes(ByVal strSub As String) As String
    
    'On Error GoTo errH
    Dim mPos1 As Integer, prevOP As Integer, nextOP As Integer
    Dim operand1 As String, operand2 As String
    Dim tmpChr As String
    Dim i As Integer, opCnt As Integer
    Dim mCpl As Complex
    Dim currOpLen As Integer, prevOpLen As Integer  'za petlje
    
    'MsgBox "IDE:" & strSub
    '
    'ako su došli parametri funkcije, onda rješavamo
    '   funkcije koje su parametri
    If InStrB(1, strSub, LNG_ARGDELIMITER) > 0 Then
        mPos1 = 0
        'ako ima koja funkcija koja je kao
        '   parametar glavne, riješimo to
        If InStrB(1, strSub, "\") <> 0 Then
            
            Do
                i = InStr(mPos1 + 1, strSub, ";")
                If i = 0 Then i = Len(strSub) + 1
                operand1 = Mid$(strSub, mPos1 + 1, i - mPos1 - 1)
                'provjerimo još jednom da li se radi o
                '   funkciji, tj. da li je trenutni argument
                '   f-ja
                If InStrB(1, operand1, "\") <> 0 Then
                    'mPos1 = mPos1 + 1
                    'If mPos1 = 0 Then mPos1 = 1
                    'MsgBox Mid$(strSub, 1, mPos1) & vbCrLf & getSubRes(operand1) & vbCrLf & Mid$(strSub, i)
                    strSub = Mid$(strSub, 1, mPos1) & getSubRes(operand1) & Mid$(strSub, i)
                End If
                mPos1 = i
            Loop While i < Len(strSub)
            
        End If
        '
        sysVars.Add strSub, LNG_SYSVAR & (sysVars.Count + 1 & LNG_DOWNLN)
        getSubRes = LNG_SYSVAR & (sysVars.Count & LNG_DOWNLN)
        Exit Function
    End If
    
    If InStrB(1, strSub, "E+", vbTextCompare) <> 0 Then
        strSub = Replace$(strSub, "E+", "E", , , vbTextCompare)
    End If
    
    opCnt = 0
    'ovdje se vrte svi operatori i provjerava
    '   se da li postoje u izrazu i ako postoje
    '   rješavaju se
    For i = 0 To opBound
    
        mPos1 = InStr(2, strSub, operators(i))
        'ako ih ima
         If mPos1 > 1 Then
            currOpLen = Len(operators(i))
            
            'ako ispred odabranog operatora nije neki
            '   drugi operator tada ga odaberemo (npr. 2*-5, 2E-3*5)
            tmpChr = Mid$(strSub, mPos1 - 1, 1)
            If InStrB(1, OP_LIST, tmpChr) = 0 Then
            
                'sve dok ima operatora koji je trenutno odabran,
                '   traže se operandi i rade operacije s njima
                Do While mPos1 > 1
                    opCnt = 1 'opCnt + 1
                    
                    'traženje operatora prije i poslije trenutnog
                    '   operatora
                    prevOP = getPrevOperator(strSub, mPos1)
                    nextOP = getNextOperator(strSub, mPos1)
                    
                    'ukoliko je uzeo predznak od prvog broja kao operator, moramo
                    '   staviti da je to 0 kako bi se i predznak uzeo
                    If prevOP = 1 Then
                        prevOP = 0
                    End If
                    
                    'ako nema sljedeceg operatora, stavimo da je to
                    '   dužina izraza kako bi se uzeo drugi operand
                    If nextOP = 0 Then
                        nextOP = Len(strSub) + 1
                    End If
                    
                    'ako je došlo do -* il tak neš što ne valja
                    If Abs(prevOP - mPos1) < 2 Or Abs(nextOP - mPos1) < 2 Then
                        lastErr = ERR_ExpectedExpres
                        lastErrNum = ERR_ExpectedExpresN
                        Exit Function
                        'Err.Raise lastErrNum, "clsCalculator", lastErr
                    End If
                    
                    'izdvajanje operanada
                    operand1 = Mid$(strSub, prevOP + 1, mPos1 - prevOP - 1)
                    operand2 = Mid$(strSub, mPos1 + currOpLen, nextOP - mPos1 - currOpLen)
                    '
                    'možda bude trebalo
'                    If InStrB(1, operand1, " ") <> 0 Then
'                        operand1 = Replace$(operand1, " ", vbNullString)
'                    End If
'                    If InStrB(1, operand2, " ") <> 0 Then
'                        operand2 = Replace$(operand2, " ", vbNullString)
'                    End If

                    'provjerimo da li je koji operand možda funkcija
                    If InStrB(1, operand1, LNG_REPLACEB) <> 0 Then
                           
                           operand1 = solveFunction(operand1)
                           
                    End If
                    If InStrB(1, operand2, LNG_REPLACEB) <> 0 Then

                            operand2 = solveFunction(operand2)

                    End If
                    
                    'MsgBox operand1 & vbCrLf & operand2
'                    If operand1 <> CPL_IMAG Then
'                        tmpChr = left$(operand1, 1)
'                        If IsNumeric(tmpChr) = False Then
'                            If InStrB(1, CHR_SOLVEFUC, tmpChr) = 0 Then 'tmpChr <> LNG_DOWNLN And tmpChr <> LNG_DOLLAR And tmpChr <> LNG_REPLACEB And tmpChr <> OP_MINUS And tmpChr <> OP_PLUS And tmpChr <> "." Then
'                                'MsgBox "CALL:" & operand1
'                                operand1 = solveFunction(operand1)
'                            End If
'                        End If
'                    End If
'                    If operand2 <> CPL_IMAG Then
'                        tmpChr = left$(operand2, 1)
'                        If IsNumeric(tmpChr) = False Then
'                            If InStrB(1, CHR_SOLVEFUC, tmpChr) = 0 Then ' tmpChr <> LNG_DOWNLN And tmpChr <> LNG_DOLLAR And tmpChr <> LNG_REPLACEB And tmpChr <> OP_MINUS And tmpChr <> OP_PLUS And tmpChr <> "." Then
'                                'MsgBox "CALL2:" & operand2
'                                operand2 = solveFunction(operand2)
'                            End If
'                        End If
'                    End If
                    'izdvoimo zagrade \
'                    If InStrB(1, operand1, LNG_REPLACEB) Then
'                        operand1 = Replace$(operand1, LNG_REPLACEB, vbNullString)
'                    End If
'                    If InStrB(1, operand2, LNG_REPLACEB) Then
'                        operand2 = Replace$(operand2, LNG_REPLACEB, vbNullString)
'                    End If
                    'ako je operand varijabla
                    'MsgBox "2:" & operand1 & vbCrLf & operand2
                    'If InStrB(1, operand1, "$") <> 0 Or InStrB(1, operand1, "_sysVar_") <> 0 Then
                    
                    'ako ima znak - ispred zagrade ili varijable onda
                    '   napravimo negaciju
                    If left$(operand1, 9) <> "-_sysVar_" And left$(operand1, 2) <> "-$" Then
                        
                        replaceVarsWithVals operand1
                        
                    Else
                    
                        operand1 = Mid$(operand1, 2)
                        replaceVarsWithVals operand1
                        operand1 = getSimleRes("-1", operand1, "*")
                        replaceVarsWithVals operand1
                        
                    End If
                    
                    '
                    If left$(operand2, 9) <> "-_sysVar_" And left$(operand2, 2) <> "-$" Then
                        
                        replaceVarsWithVals operand2
                        
                    Else
                        
                        operand2 = Mid$(operand2, 2)
                        replaceVarsWithVals operand2
                        operand2 = getSimleRes("-1", operand2, "*")
                        replaceVarsWithVals operand2
                        
                    End If
                    

                    'decimalni zarez je .
                    If InStrB(1, operand1, ",") Then
                        operand1 = Replace$(operand1, ",", ".")
                    End If
                    If InStrB(1, operand2, ",") Then
                        operand2 = Replace$(operand2, ",", ".")
                    End If
                    
                    'tražimo rezultat
                    operand1 = getSimleRes(operand1, operand2, operators(i))
                    If InStrB(1, operand1, ",") Then
                        operand1 = Replace$(operand1, ",", ".")
                    End If
                    '
                    'pronaðemo duljinu prethodnog operatora
                    prevOpLen = 1
                    If prevOP > 0 Then
                    
                        Select Case Mid$(strSub, prevOP, 1)
                            Case "!"
                                prevOpLen = 2
                                
                            Case "<"
                                If Mid$(strSub, prevOP + 1, 1) = "=" Then
                                    prevOpLen = 2
                                End If
                                
                            Case ">"
                                If Mid$(strSub, prevOP + 1, 1) = "=" Then
                                    prevOpLen = 2
                                End If
                                
                            Case "="
                                If Mid$(strSub, prevOP + 1, 1) = "=" Then
                                    prevOpLen = 2
                                End If
                                
                            Case "&"
                                prevOpLen = 2
                                
                            Case "|"
                                prevOpLen = 2
                                
                        End Select
                        
                    End If
                    
                    'izdvojimo ono što smo upravo riješili i stavimo na to mjesto
                    '   rezultat
                    If nextOP > 0 Or prevOP > 0 Then
                        
                        'provjerimo da li je duljina rezultata manja od
                        '   duljine izraza i ako jest, samo zamijenimo
                        '   to umjesto da kopiramo (radi brzine)
                        If Len(operand1) < nextOP - prevOP Then
                            
                            If Len(operand1) < nextOP - prevOP Then
                                operand1 = operand1 & String(nextOP - prevOP, " ")
                            End If
                            Mid$(strSub, prevOP + prevOpLen, nextOP - prevOP - prevOpLen) = operand1
                        
                        Else
                            
                            strSub = Mid$(strSub, 1, prevOP + prevOpLen - 1) & (operand1 & Mid$(strSub, nextOP))
                        
                        End If
                        
                    Else
                        
                        strSub = operand1
                        
                    End If
                    mPos1 = InStr(1, strSub, operators(i))
                Loop
                
            End If
            
        End If
        
    Next i
    
    
    'ukoliko nije bilo operatora, provjeravamo da
    '   li ima funkcije ili korisnicke varijable
    If opCnt = 0 Then
        
        tmpChr = Mid$(strSub, 1, 1)
        If left$(strSub, 1) = LNG_DOLLAR Then
            If checkVar(strSub) = False Then
                lastErr = ERR_UndefVar & "'" & strSub & "'"
                lastErrNum = ERR_UndefVarN
                Exit Function
            End If
            strSub = userVars(strSub)
        End If
        'ako ima samo varijablu s nekim predznakom
        If InStrB(1, strSub, "$") <> 0 Then
            If left$(strSub, 9) <> "-_sysVar_" And left$(strSub, 2) <> "-$" Then
                replaceVarsWithVals strSub
            Else
                strSub = Mid$(strSub, 2)
                replaceVarsWithVals strSub
                strSub = getSimleRes("-1", strSub, "*")
                replaceVarsWithVals strSub
            End If
        End If
        
        If IsNumeric(tmpChr) = False Then
            If InStrB(1, strSub, LNG_REPLACEB) <> 0 Then
                strSub = solveFunction(strSub)
            End If
        ElseIf InStrB(1, strSub, "\") <> 0 Then
            lastErr = ERR_ExpectedExpres
            lastErrNum = ERR_ExpectedExpresN
        End If
        
    End If
    
    
    'ako je slucajno povratak varijable, naðemo vrijednost
    If left$(strSub, 1) = "_" Then 'InStrB(1, strSub, "_sysVar_") <> 0 Then
        
        replaceVarsWithVals strSub
        If mCpl.Im <> 0 Then
            strSub = getStrCplx(mCpl)
        End If
        
    'ukoliko se nlalazi znak minus ispred varijable
    ElseIf left$(strSub, 1) = OP_MINUS And Mid$(strSub, 2, 1) = LNG_DOWNLN Then
        
        strSub = Mid$(strSub, 2)
        Do While left$(strSub, 1) = LNG_DOWNLN
            strSub = sysVars.Item(strSub)
            If InStrB(1, strSub, LNG_REPLACEB) Then
                strSub = Replace$(strSub, LNG_REPLACEB, vbNullString)
            End If
        Loop
        mCpl = getComplex(strSub)
        If mCpl.Im = 0 Then
            strSub = OP_MINUS & strSub
        Else
            mCpl.Re = -mCpl.Re
            mCpl.Im = -mCpl.Im
            strSub = getStrCplx(mCpl)
        End If
        
    End If
    
    'postavljanje vrijednosti koju vracamo
    getSubRes = strSub
    Exit Function
errH:
    lastErr = ERR_InvalidFunction
    lastErrNum = ERR_InvalidFunctionN
End Function

'##########################################################
'#      simple res    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
'##########################################################

'funkcija koja rješava jednostavan izraz od
'   dva operanda i jednog operatora
Private Function getSimleRes(ByVal operand1 As String, ByVal operand2 As String, ByVal operator As String) As String
    
    Dim cplx1 As Complex, cplx2 As Complex
    
    'brišemo zagrade sa operanada
    If InStrB(1, operand1, LNG_OPENST) Then
        operand1 = Replace$(operand1, LNG_OPENST, vbNullString)
    End If
    If InStrB(1, operand1, LNG_CLOSEST) Then
        operand1 = Replace$(operand1, LNG_CLOSEST, vbNullString)
    End If
    If InStrB(1, operand2, LNG_OPENST) Then
        operand2 = Replace$(operand2, LNG_OPENST, vbNullString)
    End If
    If InStrB(1, operand2, LNG_CLOSEST) Then
        operand2 = Replace$(operand2, LNG_CLOSEST, vbNullString)
    End If
   
    'realni i im dio
    If InStrB(1, operand1, "i") <> 0 Then
        cplx1 = getComplex(operand1)
    Else
        cplx1.Re = Val(operand1)
        cplx1.Im = 0
    End If
    If InStrB(1, operand2, "i") <> 0 Then
        cplx2 = getComplex(operand2)
    Else
        cplx2.Re = Val(operand2)
        cplx2.Im = 0
    End If
    
    '
    Select Case operator
        Case OP_POT
            
            'ako je potenciranje s 1, prekacemo
            If cplx2.Re <> 1 Or cplx2.Im <> 0 Then
                If cplx1.Im = 0 And cplx2.Im = 0 Then
                    If Abs(cplx2.Re) > 1 And cplx1.Re >= 0 Then
                        cplx1.Re = cplx1.Re ^ cplx2.Re
                    'za negativne brojeve
                    Else
                        '(-x)^y = x^y * exp(i*y*pi)
                        Dim c1 As Complex, c2 As Complex
                        c1.Re = Abs(cplx1.Re) ^ cplx2.Re
                        c1.Im = 0
                        c2.Re = 0
                        c2.Im = cplx2.Re * PI
                        c2 = CExp(c2)
                        cplx1 = CMul(c1, c2)
                    End If
                ElseIf cplx2.Im = 0 Then
                    cplx1 = CPot(cplx1, cplx2.Re)
                Else
                    cplx1 = CPot2(cplx1, cplx2)
                End If
            End If
            
        Case OP_DIV
            
            If cplx2.Im = 0 And cplx1.Im = 0 Then
                If cplx2.Re = 0 Then
                    lastErr = ERR_DivByZer
                    lastErrNum = ERR_DivByZerN
                    Exit Function
                    'Err.Raise lastErrNum, "clsCalculator", lastErr
                End If
                cplx1.Re = cplx1.Re / cplx2.Re
            ElseIf cplx2.Im = 0 Then
                If cplx2.Re = 0 Then
                    lastErr = ERR_DivByZer
                    lastErrNum = ERR_DivByZer
                    Exit Function
                    'Err.Raise lastErrNum, "clsCalculator", lastErr
                End If
                cplx1.Re = cplx1.Re / cplx2.Re
                cplx1.Im = cplx1.Im / cplx2.Re
            Else
                'MsgBox cplx1.Re & "+" & cplx1.Im & vbCrLf & cplx2.Re & "+" & cplx2.Im
                cplx1 = CDiv(cplx1, cplx2)
            End If
            
        Case OP_MUL
            
            'MsgBox cplx1.Re & OP_PLUS & cplx1.Im & vbCrLf & cplx2.Re & OP_PLUS & cplx2.Im
            If cplx2.Im = 0 And cplx1.Im = 0 Then
                cplx1.Re = cplx1.Re * cplx2.Re
            ElseIf cplx2.Im = 0 Then
                cplx1.Re = cplx1.Re * cplx2.Re
                cplx1.Im = cplx1.Im * cplx2.Re
            ElseIf cplx1.Im = 0 Then
                cplx1.Im = cplx2.Im * cplx1.Re
                cplx1.Re = cplx2.Re * cplx1.Re
            Else
                cplx1 = CMul(cplx1, cplx2)
            End If
            
        Case OP_MINUS
            
            cplx1 = CSub(cplx1, cplx2)
            
        Case OP_PLUS
            
            cplx1 = CAdd(cplx1, cplx2)
            
        'logic
        Case OP_LESSTHAN
            
            If cplx1.Im <> 0 Or cplx2.Im <> 0 Then
                If cAbs(cplx1) < cAbs(cplx2) Then
                    cplx1.Re = 1
                    cplx1.Im = 0
                Else
                    cplx1.Re = 0
                    cplx1.Im = 0
                End If
            Else
                If cplx1.Re < cplx2.Re Then
                    cplx1.Re = 1
                    cplx1.Im = 0
                Else
                    cplx1.Re = 0
                    cplx1.Im = 0
                End If
            End If
            
        Case OP_GREATERTHAN
            
            If cplx1.Im <> 0 Or cplx2.Im <> 0 Then
                If cAbs(cplx1) > cAbs(cplx2) Then
                    cplx1.Re = 1
                    cplx1.Im = 0
                Else
                    cplx1.Re = 0
                    cplx1.Im = 0
                End If
            Else
                If cplx1.Re > cplx2.Re Then
                    cplx1.Re = 1
                    cplx1.Im = 0
                Else
                    cplx1.Re = 0
                    cplx1.Im = 0
                End If
            End If
            
        Case OP_GREATERORIS
            
            If cplx1.Im <> 0 Or cplx2.Im <> 0 Then
                If cAbs(cplx1) >= cAbs(cplx2) Then
                    cplx1.Re = 1
                    cplx1.Im = 0
                Else
                    cplx1.Re = 0
                    cplx1.Im = 0
                End If
            Else
                If cplx1.Re >= cplx2.Re Then
                    cplx1.Re = 1
                    cplx1.Im = 0
                Else
                    cplx1.Re = 0
                    cplx1.Im = 0
                End If
            End If
            
        Case OP_LESSORIS
            
            If cplx1.Im <> 0 Or cplx2.Im <> 0 Then
                If cAbs(cplx1) <= cAbs(cplx2) Then
                    cplx1.Re = 1
                    cplx1.Im = 0
                Else
                    cplx1.Re = 0
                    cplx1.Im = 0
                End If
            Else
                If cplx1.Re <= cplx2.Re Then
                    cplx1.Re = 1
                    cplx1.Im = 0
                Else
                    cplx1.Re = 0
                    cplx1.Im = 0
                End If
            End If
            
        Case OP_ISNOT
            
            If cplx1.Im <> 0 Or cplx2.Im <> 0 Then
                If cAbs(cplx1) <> cAbs(cplx2) Then
                    cplx1.Re = 1
                    cplx1.Im = 0
                Else
                    cplx1.Re = 0
                    cplx1.Im = 0
                End If
            Else
                If cplx1.Re <> cplx2.Re Then
                    cplx1.Re = 1
                    cplx1.Im = 0
                Else
                    cplx1.Re = 0
                    cplx1.Im = 0
                End If
            End If
            
        Case OP_IS
            '
            If cplx1.Im <> 0 Or cplx2.Im <> 0 Then
                If cplx1.Re = cplx2.Re Then
                    If cplx1.Im = cplx2.Im Then
                        cplx1.Re = 1
                    Else
                        cplx1.Re = 0
                    End If
                Else
                    cplx1.Re = 0
                End If
                cplx1.Im = 0
            Else
                If cplx1.Re = cplx2.Re Then
                    cplx1.Re = 1
                    cplx1.Im = 0
                Else
                    cplx1.Re = 0
                    cplx1.Im = 0
                End If
            End If
            
        Case OP_AND
            
            If cplx1.Im <> 0 Or cplx2.Im <> 0 Then
                cplx1.Re = cAbs(cplx1) And cAbs(cplx2)
                cplx1.Im = 0
            Else
                cplx1.Re = cplx1.Re And cplx2.Re
                cplx1.Im = 0
            End If
            
        Case OP_OR
        
            If cAbs(cplx1) And cAbs(cplx2) Then
                cplx1.Re = cAbs(cplx1) Or cAbs(cplx2)
                cplx1.Im = 0
            Else
                cplx1.Re = cplx1.Re Or cplx2.Re
                cplx1.Im = 0
            End If
            
    End Select
    
    
    If cplx1.Im <> 0 Then
        sysVars.Add getStrCplx(cplx1), LNG_SYSVAR & (sysVars.Count + 1 & LNG_DOWNLN)
        getSimleRes = LNG_SYSVAR & (sysVars.Count & LNG_DOWNLN)
    Else
        getSimleRes = cplx1.Re
    End If

End Function
'
'/////////////////////////////////////////////////////////////
'                   RJEŠAVANJE FUNKCIJA
'/////////////////////////////////////////////////////////////
'funkcija koja izdvaja argumente tražene funkcije
'   i šalje ih na obradu kako bi se dobio
'   rezultat
Private Function solveFunction(ByVal strFunc As String) As String

    Dim fName As String, fParams() As String, tmpParams As String
    Dim mPos1 As Integer, mPos2 As Integer, paramsCnt As Integer
    Dim i As Integer, bnd As Integer
    Dim isSpecFunc As Boolean
    
    'treba poslati i varijable clsFunctions koja ce ih vratiti svom clsCalc
    'ukoliko je došla f-ja oblika fnam_sysVar_n
    If InStrB(1, strFunc, LNG_REPLACEB) = 0 Then
    
        strFunc = Replace$(strFunc, "_sys", "\_sys")
        strFunc = strFunc & LNG_REPLACEB
        
    End If
    
    
    mPos1 = InStr(1, strFunc, LNG_REPLACEB)
    
    If mPos1 > 0 Then
        
        fName = Mid$(strFunc, 1, mPos1 - 1)
        
        'ukoliko je funkcija sum onda to posebno radimo
        '   jer se ne rješava izraz
        If LCase(fName) = "sum" Then
            isSpecFunc = True
        Else
            isSpecFunc = False
        End If
        
         'uèitamo parametre
        tmpParams = Replace$(Mid$(strFunc, mPos1 + 1), LNG_REPLACEB, vbNullString)
        'tmpParams = Replace$(tmpParams, LNG_REPLACEB, vbNullString)
        
        'za specFunc se mijenjaju samo sistemske varijable
        replaceVarsWithVals tmpParams, True, isSpecFunc
        
        'broj parametara
        bnd = getInstrCnt(tmpParams, ";") '+ 1
        ReDim fParams(bnd)
        paramsCnt = 0
        
        'tražimo sve parametre
        mPos1 = InStr(1, tmpParams, LNG_ARGDELIMITER)
        If mPos1 > 0 Then
        
            'izdvoimo vrijednost parametra i ako je izraz
            '   šaljemo ga na obradu
            fParams(paramsCnt) = Mid$(tmpParams, 1, mPos1 - 1)
            '
            
            If IsNumeric(fParams(paramsCnt)) <> True Then
            
                If isSpecFunc = False Then fParams(paramsCnt) = Calculate(fParams(paramsCnt))
            
            End If
            paramsCnt = paramsCnt + 1
            
            'ostali èlanovi
            Do While mPos1 > 0
            
                mPos2 = InStr(mPos1 + 1, tmpParams, LNG_ARGDELIMITER)
                If mPos2 > 0 Then
                    fParams(paramsCnt) = Mid$(tmpParams, mPos1 + 1, mPos2 - mPos1 - 1)
                Else
                    fParams(paramsCnt) = Mid$(tmpParams, mPos1 + 1)
                    
                    
                    If InStrB(1, fParams(paramsCnt), LNG_OPENST) = 0 Then
                        
                        fParams(paramsCnt) = Replace$(fParams(paramsCnt), LNG_CLOSEST, vbNullString)
                    
                    End If
                End If
                
                'ako je izraz, šaljemo ga na obradu
                If IsNumeric(fParams(paramsCnt)) <> True Then
                    
                    'If isSpecFunc = False Then
                    fParams(paramsCnt) = Calculate(fParams(paramsCnt))
                
                End If
                'fParams(paramsCnt) = Calculate(fParams(paramsCnt))
                'fParams(paramsCnt) = Replace$(fParams(paramsCnt), LNG_OPENST, VBNullstring)
                mPos1 = mPos2
                paramsCnt = paramsCnt + 1
                
            Loop
        'izdvajanje jedinog èlana
        Else
            
            If IsNumeric(fParams(0)) <> True Then
                fParams(0) = Calculate(tmpParams)
            End If
            paramsCnt = 1
            
        End If
    Else
        
        lastErr = ERR_ExpectArgs & strFunc
        lastErrNum = ERR_ExpectArgsN
        Exit Function
        
    End If
    '
    
    'ukoliko je koji parametar varijabla, naðemo vrijednost
    'bnd = UBound(fParams)
    For i = 0 To bnd

        Do While Mid$(fParams(i), 1, 1) = LNG_DOWNLN
            fParams(i) = Replace$(sysVars.Item(fParams(i)), LNG_REPLACEB, vbNullString)
        Loop
        
        'MsgBox fParams(i)
        'ukoliko je param oblika fname_sysVar_n_
        'replaceVarsWithVals fParams(i)         '<<<<<<<<<<<<<
        
    Next i
    
    'ukoliko nije f-ja sume onda šaljemo na obradu
    If isSpecFunc = False Then
    
    solveFunction = mFunc.getFuncResult(fName, fParams)
    
    Else
        
        solveFunction = getSum(fParams)
        
    End If

    Erase fParams
End Function

'funkcija daje sumu funkcije args(0) za $x od
'   args(1) do args(2)
Private Function getSum(ByRef args() As String, Optional isMul As Boolean = False) As String

    Dim i As Long, lwr As Long, upr As Long
    Dim res1 As Complex

    'provjera broja argumenata
    If UBound(args) < 2 Or UBound(args) > 2 Then

        lastErr = ERR_UndefFunc & ": 'sum' or wrong number of arguments!"
        lastErrNum = ERR_UndefFuncN
        getSum = vbNullString
        Exit Function

    End If


    lwr = CLng(Val(args(1)))
    upr = CLng(Val(args(2)))

    'provjera donje i gornje granice
    If lwr > upr Then

        lastErr = ERR_SumArg
        lastErrNum = ERR_SumArgN
        getSum = vbNullString
        Exit Function

    End If


    
    
    'zbrajanje
    If isMul = False Then
        res1.Re = 0
        res1.Im = 0
        For i = lwr To upr
    
            Me.defineVariable "$x", Str$(i)
            res1 = CAdd(res1, getComplex(Me.Calculate(args(0))))
    
        Next i
    Else
        res1.Re = 1
        res1.Im = 1
        For i = lwr To upr
    
            Me.defineVariable "$x", Str$(i)
            res1 = CMul(res1, getComplex(Me.Calculate(args(0))))
    
        Next i
    End If

    getSum = Replace$(getStrCplx(res1), ",", ".")
    
End Function
'
'izdvajanje varijabli iz stringa i zamjenja
'   njih sa njihovim vrijednostima
'
Private Sub replaceVarsWithVals(ByRef mFnc As String, Optional solveFunc As Boolean = False, Optional onlySys As Boolean = False)
    
    Dim mPos1 As Integer, mPos2 As Integer
    Dim mVar As String, tmpPos As Integer
    'zamjena systemskih varijabli sa
    '   njihovim vrijednostima
    mPos1 = InStr(1, mFnc, LNG_SYSVAR)
    
    Do While mPos1 > 0
    
        mPos2 = InStr(mPos1 + Len(LNG_SYSVAR) + 1, mFnc, LNG_DOWNLN)
        mVar = Trim$(Mid$(mFnc, mPos1, mPos2 - mPos1 + 1))
        Do While left$(mVar, 1) = LNG_DOWNLN
            'ako je vrijednost kompleksni broj, stavljamo ga u zagrade
            If InStrB(1, mVar, ";") Then
                replaceVarsWithVals mVar
                Exit Do
            End If
            '
'            If InStrB(1, sysVars.Item(mVar), CPL_IMAG) <> 0 And InStrB(1, sysVars.Item(mVar), LNG_REPLACEB) = 0 And InStrB(1, sysVars.Item(mVar), ";") = 0 Then
'                mVar = LNG_REPLACEB & sysVars.Item(mVar) & LNG_REPLACEB
'            Else
                mVar = sysVars.Item(mVar)
                
'            End If
            'brišemo \
            If InStrB(1, mVar, LNG_REPLACEB) <> 0 Then
                mVar = Replace$(mVar, LNG_REPLACEB, vbNullString)
            End If
            
        Loop
        mFnc = Mid$(mFnc, 1, mPos1 - 1) & (mVar & Mid$(mFnc, mPos2 + 1))
        mPos1 = InStr(1, mFnc, LNG_SYSVAR)
    Loop
    '
    
    'ako je zamjena samo sistemskih
    If onlySys = True Then
        
        Exit Sub
        
    End If
    
    'zamjena korisnickih varijabli sa
    '   njihovim vrijednostima
    '
    mPos1 = InStr(1, mFnc, LNG_DOLLAR)
    Do While mPos1 > 0
        
        mPos2 = getNextOperator(mFnc, mPos1)
        
        'provjera da li je možda otvorena zagrada bliže
        tmpPos = InStr(mPos1, mFnc, LNG_CLOSEST)
        
        If tmpPos > 0 Then
            
            If tmpPos < mPos2 Then
                mPos2 = tmpPos
            End If
            
        End If
        mPos2 = mPos2 - 1
        
        'ako je mPos2 < 1
        If mPos2 < 1 Then
            mPos2 = InStr(mPos1, mFnc, ")")
            If mPos2 < 1 Then mPos2 = Len(mFnc)
        End If
        
        'provjera da li je ; prije operanda ili karaja
        tmpPos = InStr(mPos1, mFnc, ";")
        
        If tmpPos < mPos2 Then
            
            If tmpPos > 0 Then
                mPos2 = tmpPos - 1
            End If
            
        End If
        '
        mVar = Mid$(mFnc, mPos1, mPos2 - mPos1 + 1)
        Do While Mid$(mVar, 1, 1) = LNG_DOLLAR
           
            mVar = Replace$(mVar, LNG_OPENST, vbNullString)
            mVar = Replace$(mVar, LNG_CLOSEST, vbNullString)

            'varijabla ne postoji
            If checkVar(mVar) <> True Then
                
                'ukoliko nije argument f-je zato jer
                '   je moguce imati izraz kao argument
                '   a ako varijabla u izrazu ipak ne postoji,
                '   skužit ce kod rješavanja izraza
                If solveFunc <> True Then
                
                    lastErr = ERR_UndefVar & "'" & mVar & "'"
                    lastErrNum = ERR_UndefVarN
                    
                End If
                
                Exit Sub
                
            End If
            
            
            'MsgBox mVar

            'ako je vrijednost kompleksni broj u zagrade
            If InStrB(1, userVars.Item(mVar), CPL_IMAG) > 0 Then
                mVar = LNG_OPENST & userVars.Item(mVar) & LNG_CLOSEST
                
            Else
                mVar = userVars.Item(mVar)
                
            End If

            mVar = Replace$(mVar, LNG_REPLACEB, vbNullString)

            
        Loop
        
        mFnc = Mid$(mFnc, 1, mPos1 - 1) & (mVar & Mid$(mFnc, mPos2 + 1))
        mPos1 = InStr(mPos1 + 1, mFnc, LNG_DOLLAR)
        
    Loop
End Sub

'provjera da li je varijabla definirana
Private Function checkVar(ByVal varName As String) As Boolean
    
    On Error Resume Next
    Dim k As String
    k = userVars(varName)
    If Err.Number = 0 Then
        checkVar = True
    Else
        checkVar = False
    End If
    
End Function
'
Private Sub Class_Initialize()

    ReDim operators(12)
    operators(0) = OP_POT
    operators(1) = OP_DIV
    operators(2) = OP_MUL
    '
    operators(3) = OP_MINUS 'OP_AND
    operators(4) = OP_PLUS ' OP_OR
    '
    'operatori usporedbe
    operators(5) = OP_LESSORIS
    operators(6) = OP_GREATERORIS
    operators(7) = OP_LESSTHAN
    operators(8) = OP_GREATERTHAN
    operators(9) = OP_ISNOT
    operators(10) = OP_IS
    'logicki operatori
    operators(11) = OP_AND
    operators(12) = OP_OR
    
End Sub

Private Sub Class_Terminate()

    Set sysVars = Nothing
    Set userVars = Nothing
    Set userVarNames = Nothing
    '
    Set mFunc = Nothing
    Erase operators
    
End Sub
