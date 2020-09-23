Attribute VB_Name = "modString"
Option Explicit
'Private Const OP_LIST   As String = "()+-*/%^<=>=!=&&||"
Private Const OP_BEF   As String = "+-*/%^<=>=!=&&||Ee" '"()eEi"
Private Const LGOP_BEF   As String = "<>!"
'broj stringova lfStr u  liStr
Public Function getInstrCnt(ByVal liStr As String, ByVal lfStr As String, Optional caseSens As Boolean = True) As Integer
    Dim pos1 As Integer
    getInstrCnt = 0
    
    If caseSens = False Then
        pos1 = InStr(1, liStr, lfStr)
    Else
        pos1 = InStr(1, liStr, lfStr, vbTextCompare)
    End If
    
    If pos1 > 0 Then
        Do While pos1 > 0
            getInstrCnt = getInstrCnt + 1
            If caseSens = False Then
                pos1 = InStr(pos1 + 1, liStr, lfStr)
            Else
                pos1 = InStr(pos1 + 1, liStr, lfStr, vbTextCompare)
            End If
        Loop
    End If
End Function
'traženje operatora prije trenutnog
Public Function getPrevOperator(ByVal mStr As String, ByVal mPos As Integer) As Integer
    If mPos <= 1 Then
        getPrevOperator = 0
        Exit Function
    End If
    Dim pos(12) As Integer, i As Integer, strPrev As String
    pos(0) = InStrRev(mStr, OP_POT, mPos - 1)
    pos(1) = InStrRev(mStr, OP_DIV, mPos - 1)
    pos(2) = InStrRev(mStr, OP_MUL, mPos - 1)
    pos(3) = InStrRev(mStr, OP_MINUS, mPos - 1)
    
    pos(5) = InStrRev(mStr, OP_GREATERORIS, mPos - 1)
    pos(6) = InStrRev(mStr, OP_LESSORIS, mPos - 1)
    pos(7) = InStrRev(mStr, OP_GREATERTHAN, mPos - 1)
    pos(8) = InStrRev(mStr, OP_LESSTHAN, mPos - 1)
    
    pos(9) = InStrRev(mStr, OP_ISNOT, mPos - 1)
    'pos(10) = InStrRev(mStr, OP_IS, mPos - 1)
    
    pos(11) = InStrRev(mStr, OP_AND, mPos - 1)
    pos(12) = InStrRev(mStr, OP_OR, mPos - 1)
   ' pos(5) = InStrRev(mStr, ",", mPos - 1)
    
    If pos(3) > 1 Then
        strPrev = Mid$(mStr, pos(3) - 1, 1)
        Do While InStrB(1, OP_BEF, strPrev) <> 0 'IsNumeric(strPrev) <> True ' StrComp(strPrev, "E", vbTextCompare) = 0 Or strPrev = OP_POT Or strPrev = OP_MUL Or strPrev = OP_DIV Or strPrev = OP_PLUS Or strPrev = OP_MINUS
            pos(3) = InStrRev(mStr, OP_MINUS, pos(3) - 1)
            If pos(3) < 2 Then
                pos(3) = 0
                Exit Do
            End If
            strPrev = Mid$(mStr, pos(3) - 1, 1)
        Loop
    End If
    
    pos(4) = InStrRev(mStr, OP_PLUS, mPos - 1)
    If pos(4) > 1 Then
        strPrev = Mid$(mStr, pos(4) - 1, 1)
        Do While InStrB(1, OP_BEF, strPrev) <> 0 'IsNumeric(strPrev) <> True 'StrComp(strPrev, "E", vbTextCompare) = 0 Or strPrev = OP_POT Or strPrev = OP_MUL Or strPrev = OP_DIV Or strPrev = OP_PLUS Or strPrev = OP_MINUS
            pos(4) = InStrRev(mStr, OP_PLUS, pos(4) - 1)
            If pos(4) < 2 Then
                pos(4) = 0
                Exit Do
            End If
            strPrev = Mid$(mStr, pos(4) - 1, 1)
        Loop
    End If
    
    pos(10) = InStrRev(mStr, OP_IS, mPos - 1)
    If pos(10) > 1 Then
        strPrev = Mid$(mStr, pos(10) - 1, 1)
        Do While InStrB(1, LGOP_BEF, strPrev) <> 0 'IsNumeric(strPrev) <> True 'StrComp(strPrev, "E", vbTextCompare) = 0 Or strPrev = OP_LESSTHAN Or strPrev = OP_GREATERTHAN Or strPrev = "!"
            pos(10) = InStrRev(mStr, OP_IS, pos(10) - 1)
            If pos(10) < 2 Then
                pos(10) = 0
                Exit Do
            End If
            strPrev = Mid$(mStr, pos(10) - 1, 1)
        Loop
    End If
    getPrevOperator = getMinDist(pos, mPos)
    '
End Function
'traženje sljedeæeg operatora
Public Function getNextOperator(ByVal mStr As String, ByVal mPos As Integer) As Integer
    If mPos <= 1 Then
        getNextOperator = 0
        Exit Function
    End If
    Dim pos(12) As Integer, i As Integer, strPrev As String
    
    pos(0) = InStr(mPos + 1, mStr, OP_POT)
    pos(1) = InStr(mPos + 1, mStr, OP_DIV)
    pos(2) = InStr(mPos + 1, mStr, OP_MUL)
    pos(3) = InStr(mPos + 1, mStr, OP_MINUS)
    
    pos(5) = InStr(mPos + 1, mStr, OP_GREATERORIS)
    pos(6) = InStr(mPos + 1, mStr, OP_LESSORIS)
    pos(7) = InStr(mPos + 1, mStr, OP_GREATERTHAN)
    pos(8) = InStr(mPos + 1, mStr, OP_LESSTHAN)
    
    pos(9) = InStr(mPos + 1, mStr, OP_ISNOT)
    
    pos(11) = InStr(mPos + 1, mStr, OP_AND)
    pos(12) = InStr(mPos + 1, mStr, OP_OR)
    
    '-
    If pos(3) > 1 Then
        strPrev = Mid$(mStr, pos(3) - 1, 1)
        If strPrev <> "\" Then
            Do While InStrB(1, OP_BEF, strPrev) <> 0 ' IsNumeric(strPrev) <> True  'StrComp(strPrev, "E", vbTextCompare) = 0 Or strPrev = OP_POT Or strPrev = OP_MUL Or strPrev = OP_DIV Or strPrev = OP_PLUS Or strPrev = OP_MINUS
                pos(3) = InStr(pos(3) + 1, mStr, OP_MINUS)
                If pos(3) < 2 Then
                    pos(3) = 0
                    Exit Do
                End If
                strPrev = Mid$(mStr, pos(3) - 1, 1)
            Loop
        End If
    End If
    
    '+
    pos(4) = InStr(mPos + 1, mStr, OP_PLUS)
    If pos(4) > 1 Then
        strPrev = Mid$(mStr, pos(4) - 1, 1)
        If strPrev <> "\" Then
            Do While InStrB(1, OP_BEF, strPrev) <> 0 ' IsNumeric(strPrev) <> True 'StrComp(strPrev, "E", vbTextCompare) = 0 Or strPrev = OP_POT Or strPrev = OP_MUL Or strPrev = OP_DIV Or strPrev = OP_PLUS Or strPrev = OP_MINUS
                pos(4) = InStr(pos(4) + 1, mStr, OP_PLUS)
                If pos(4) < 2 Then
                    pos(4) = 0
                    Exit Do
                End If
                strPrev = Mid$(mStr, pos(4) - 1, 1)
            Loop
        End If
    End If
    '==
    pos(10) = InStr(mPos + 1, mStr, OP_IS)
    If pos(10) > 1 Then
        strPrev = Mid$(mStr, pos(10) - 1, 1)
        If strPrev <> "\" Then
            Do While InStrB(1, LGOP_BEF, strPrev) <> 0 ' IsNumeric(strPrev) <> True 'StrComp(strPrev, "E", vbTextCompare) = 0 Or strPrev = OP_LESSTHAN Or strPrev = OP_GREATERTHAN Or strPrev = "!"
                pos(10) = InStr(pos(10) + 1, mStr, OP_PLUS)
                If pos(10) < 2 Then
                    pos(10) = 0
                    Exit Do
                End If
                strPrev = Mid$(mStr, pos(10) - 1, 1)
            Loop
        End If
    End If
    getNextOperator = getMinDist(pos, mPos)
End Function
'traženje najmanjeg razmaka izmeðu dva operatora
Public Function getMinDist(ByRef mArr() As Integer, ByVal pos1 As Integer) As Integer
    Dim i As Integer ', min As Integer
    Dim bnd As Integer
    bnd = UBound(mArr)
    
    For i = 0 To bnd
        If mArr(i) > 0 Then
            getMinDist = mArr(i)
            Exit For
        End If
    Next i
    For i = i To bnd
        If Abs(pos1 - mArr(i)) < Abs(pos1 - getMinDist) Then
            If Abs(pos1 - mArr(i)) > 0 And mArr(i) > 0 Then
                getMinDist = mArr(i)
            End If
        End If
    Next i
    'if there is something like 5+(-6), for program it's 5+-6 and we need to use +
'    For i = 0 To bnd
'        If min - mArr(i) = 1 Then
'            min = mArr(i)
'        End If
'    Next i
   ' getMinDist = min
End Function
'od fname(arg1;arg2...argN): fname(N)
Public Function getFuncName(ByVal defString As String) As String
    'fname(par1;par2...)
    Dim i As Integer, num1 As Integer, end1 As Integer, start1 As String
    Dim fName As String
    'brisanje praznina
    If InStrB(1, defString, " ") <> 0 Then
        defString = Replace$(defString, " ", vbNullString)
    End If
    i = InStr(1, defString, "(")
    fName = Mid$(defString, LEN_FUNCTION + 1, i - LEN_FUNCTION - 1)
    num1 = 0
    start1 = i
    num1 = getInstrCnt(defString, LNG_ARGDELIMITER)
    'if there is one on there is no params
    If num1 = 0 Then
        end1 = InStr(i, defString, ")")
        If end1 - start1 > 1 Then num1 = 1
    Else
        num1 = num1 + 1
    End If
    getFuncName = fName & Replace$("(N)", "N", num1) '("(" & num1 & ")")
End Function
'izdvaja varijable iz definicje f-je i postavlja ih u kolekcije,
'   u jednu kolekciju stavlja 0, naziv
'   u drugu naziv
'Public Sub populateWithArgVars(ByVal strF As String, ByRef mColNames As Collection)
'    Dim i As Integer, end1 As Integer, start1 As String ',num1 as integer
'    Dim tmpArr() As String
'    'Dim tmpArgs As String
'
'    start1 = Instr$(1, strF, "(")
'    end1 = InStrRev(strF, ")")
'    tmpArr = Split(Mid$(strF, start1 + 1, end1 - start1 - 1), ";")
'    For i = 0 To UBound(tmpArr)
'        mColNames.Add tmpArr(i)
'    Next i
'    Erase tmpArr
'
'
''    'brisanje praznina
''    If InStrB(1, strF, " ") <> 0 Then
''        strF = Replace$(strF, " ", vbNullString)
''    End If
''    i = Instr$(1, strF, "(")
''    fName = Mid$(strF, LEN_FUNCTION + 1, i - LEN_FUNCTION - 1)
''    num1 = 0
''    start1 = i
''    num1 = getInstrCnt(strF, LNG_ARGDELIMITER)
''    'MsgBox "IDE1:" & strF
''    'if there is one on there is no params
''    If num1 = 0 Then
''        end1 = Instr$(i, strF, ")")
''        If end1 - start1 > 1 Then
''            start1 = Instr$(1, strF, "(")
''            'mCol.Add 0, Mid$(strF, start1 + 1, end1 - start1 - 1)
''            mColNames.Add Mid$(strF, start1 + 1, end1 - start1 - 1)
''            'MsgBox mid$(strF, start1 + 1, end1 - start1 - 1)
''        End If
''    Else
''        start1 = Instr$(1, strF, "(")
''        i = Instr$(start1, strF, LNG_ARGDELIMITER)
''        Do While i > 0
''            end1 = i
''            'MsgBox Mid$(strF, start1 + 1, end1 - start1 - 1) & vbCrLf & Mid$(strF, start1 + 1, end1 - start1 - 1)
''            'mCol.Add 0, Mid$(strF, start1 + 1, end1 - start1 - 1)
''            mColNames.Add Mid$(strF, start1 + 1, end1 - start1 - 1)
''            i = Instr$(i + 1, strF, LNG_ARGDELIMITER)
''            start1 = end1
''            If i = 0 Then
''                end1 = Len(strF)
''                'mCol.Add 0, Mid$(strF, start1 + 1, end1 - start1 - 1)
''                mColNames.Add Mid$(strF, start1 + 1, end1 - start1 - 1)
''                Exit Do
''            End If
''        Loop
''    End If
'End Sub
'briše praznine sa poèetka i kraja stringa
Public Sub ClearString(ByRef mStr As String)
    'MsgBox "D:" & mStr
    
    If InStrB(1, mStr, vbCrLf) Then
        mStr = Replace$(mStr, vbCrLf, vbNullString)
    End If
    '
    If InStrB(1, mStr, Chr$(9)) Then
        mStr = Replace$(mStr, Chr$(9), vbNullString)
    End If
    '
    If InStrB(1, mStr, " ") Then
        mStr = Trim$(mStr)
    End If
    '
    If InStrB(1, mStr, "  ") Then
        mStr = Replace$(mStr, "  ", " ")
    End If
    '
    
    'MsgBox "O:" & mStr
End Sub

'provjeravanje ispravnosti unesene f-je
'   vraæa 0 ako je sve ok, ako je greška vraca broj <> 0, najcešce poziciju greške
Public Function checkFunction(ByVal mFunc As String) As Integer
    If getInstrCnt(mFunc, "(") <> getInstrCnt(mFunc, ")") Then
        checkFunction = 1
    ElseIf InStrB(1, mFunc, "\") > 0 Then
        checkFunction = 2
    'ElseIf InStrB(1, mFunc, "_") > 0 Then
        'checkFunction = 2
    ElseIf InStrB(1, mFunc, ":") > 0 Then
        checkFunction = 2
    ElseIf InStrB(1, mFunc, ")(") > 0 Then
        checkFunction = 3
    End If
End Function

'zamjena naziva varijable sa vrijednosti
'Public Sub replaceVarVal(ByRef strF As String, ByVal strVar As String, ByVal strVal As String)
'    Dim mPos1 As Integer, strChr1 As String, strChr2 As String
'    Dim mLen As Integer
'    mLen = Len(strVar)
'
'    If InStrB(1, strF, " ") <> 0 Then
'        strF = Replace$(strF, " ", vbNullString)
'    End If
'    mPos1 = Instr$(1, strF, strVar)
'    '
'    'MsgBox "REP:" & strF & vbCrLf & strVar
'    Do While mPos1 > 0
'        'provjeravamo što se nalazi ispre i iza pronaðenog
'        If mPos1 > 1 Then
'            strChr1 = Mid$(strF, mPos1 - 1, 1)
'        Else
'            strChr1 = OP_MINUS
'        End If
'        'MsgBox strF
'        strChr2 = Mid$(strF, mPos1 + mLen, 1)
'        If Len(strChr2) = 0 Then
'            strChr2 = OP_MINUS
'        End If
'        'provjeravamo da li se radi o varijabli ili necemu drugome
'        If StrComp(strVar, Mid$(strF, mPos1, mLen), vbTextCompare) = 0 Then
'            'MsgBox mid$(strF, 1, mPos1)
'            If InStrB(1, OP_LIST, strChr1) <> 0 Then 'strChr1= "(" Or strChr1 = OP_PLUS Or strChr1 = OP_MINUS Or strChr1 = OP_DIV Or strChr1 = OP_MUL Or strChr1 = OP_POT Or strChr1 = LNG_ARGDELIMITER Or strChr1 = OP_GREATERTHAN Or strChr1 = OP_LESSTHAN Or strChr1 = OP_IS Or strChr1 = "!" Or strChr1 = "&" Or strChr1 = "|" Then
'                'MsgBox "PROLAZI1"
'                If InStrB(1, OP_LIST, strChr2) <> 0 Then  'strChr2 = ")" Or strChr2 = OP_PLUS Or strChr2 = OP_MINUS Or strChr2 = OP_DIV Or strChr2 = OP_MUL Or strChr2 = OP_POT Or strChr2 = LNG_ARGDELIMITER Or strChr2 = OP_GREATERTHAN Or strChr2 = OP_LESSTHAN Or strChr2 = OP_IS Or strChr2 = "!" Or strChr2 = "&" Or strChr2 = "|" Then
'                    'MsgBox "PROLAZI2"
'                    strF = Mid$(strF, 1, mPos1 - 1) & (strVal & Mid$(strF, mPos1 + mLen))
'                    'MsgBox mid$(strF, 1, mPos1 - 1) & vbCrLf & mid$(strF, mPos1 + len$(strVar))
'                End If
'            End If
'        End If
'        mPos1 = Instr$(mPos1 + 1, strF, strVar)
'    Loop
'    'replaceVarVal = strF
'    'MsgBox "ret =" & replaceVarVal
'End Sub
'ako dobijemo ((2+5 = 2+5
'Public Function repairVar(ByVal varLn As String) As String
'    Dim cnt1 As Integer, cnt2 As Integer
'    Dim pos1 As Integer, pos2 As Integer, lvl1 As Integer, i As Integer
'    Dim st As Boolean
'    cnt1 = getInstrCnt(varLn, "(")
'    cnt2 = getInstrCnt(varLn, ")")
'    'MsgBox varLn & vbctlf & cnt1 & OP_MINUS & cnt2
'    st = False
'    If cnt1 <> cnt2 Then
'        If cnt1 = 0 Or cnt2 = 0 Then
'            varLn = Replace$(varLn, "(", vbNullString)
'            varLn = Replace$(varLn, ")", vbNullString)
'        ElseIf cnt1 > cnt2 Then
'            lvl1 = 0
'            For i = Len(varLn) To 1 Step -1
'                If Mid$(varLn, i, 1) = ")" Then
'                    lvl1 = lvl1 + 1
'                    st = True
'                ElseIf Mid$(varLn, i, 1) = "(" Then
'                    lvl1 = lvl1 - 1
'                End If
'                '
'                If lvl1 < 0 And st = True Then
'                    repairVar = Mid$(varLn, i + 1)
'                    Exit Function
'                End If
'            Next i
'        Else
'            lvl1 = 0
'            For i = 1 To Len(varLn)
'                If Mid$(varLn, i, 1) = "(" Then
'                    lvl1 = lvl1 + 1
'                    st = True
'                ElseIf Mid$(varLn, i, 1) = ")" Then
'                    lvl1 = lvl1 - 1
'                End If
'                '
'                If lvl1 < 0 And st = True Then
'                    repairVar = Mid$(varLn, 1, i - 1)
'                    Exit Function
'                End If
'            Next i
'        End If
'    End If
'    repairVar = varLn
'End Function

'repairFunction
'Public Sub repairFunction(ByRef mFnc As String)
'    mFnc = Replace$(mFnc, "++", OP_PLUS)
'    mFnc = Replace$(mFnc, "--", OP_PLUS)
'    mFnc = Replace$(mFnc, "+-", OP_MINUS)
'    mFnc = Replace$(mFnc, "-+", OP_MINUS)
'    If Mid$(mFnc, 1, 1) = OP_PLUS Then mFnc = Mid$(mFnc, 2)
'    'MsgBox "VRAC1:" & mFnc
'    'repairFunction = mFnc
'End Sub
