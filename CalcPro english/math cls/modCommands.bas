Attribute VB_Name = "modCommands"
Option Explicit
Public lastErr As String
Public lastErrNum As Integer
'Public isResInf As Boolean

'///////////////////////////////////////////////////////////////////////
'               define
'///////////////////////////////////////////////////////////////////////
Public Sub getDefineVars(ByVal defineLine As String, ByRef retVars() As String, ByRef retVals() As String)
    
    Dim mPos1 As Integer, mPos2 As Integer, varNum As Integer, k As Integer
    Dim mVr As String, errDesc As String
    'tražimo broj varijabli
    varNum = getInstrCnt(defineLine, ",") + 1
    ReDim retVars(varNum - 1)
    ReDim retVals(varNum - 1)
    'ako ima neka greška
'    If checkDefine(defineLine, errDesc) = False Then
'        Err.Raise -11, "clsCalculator", "Sintax error: " & vbCrLf & errDesc & defineLine & vbCrLf & vbCrLf & "FILE:" & vbCrLf & FileName
'    End If
    '
    If InStrB(1, defineLine, LNG_ARGDELIMITER) <> 0 Then
    
        defineLine = Replace$(defineLine, LNG_ARGDELIMITER, vbNullString)
        
    End If
    '
    k = 0
    If varNum > 0 Then
    
        mPos1 = InStr(1, defineLine, "define", vbTextCompare) + LEN_DEFINE - 1
        Do While mPos1 > 0
            
            mPos2 = InStr(mPos1 + 1, defineLine, ",")
            If mPos2 = 0 Then mPos2 = Len(defineLine)
            mVr = Trim$(Replace$(Mid$(defineLine, mPos1 + 1, mPos2 - mPos1), ",", vbNullString))
            
            'ako ima vrijednost
            If InStr(1, mVr, "=") > 0 Then
                
                retVars(k) = Trim$(Mid$(mVr, 1, InStr(1, mVr, "=") - 1))
                retVals(k) = Trim$(Mid$(mVr, InStr(1, mVr, "=") + 1))
                
                If checkVarName(retVars(k)) = False Then
                    Err.Raise -11, "clsCalculator", "Invalid variable '" & mVr & "'" & vbCrLf & vbCrLf '& "File:" & vbCrLf & FileName
                End If
                k = k + 1
                
            Else
                
                retVars(k) = Trim$(mVr)
                If checkVarName(retVars(k)) = False Then
                    Err.Raise -11, "clsCalculator", "Invalid variable '" & mVr & "'" & vbCrLf & vbCrLf '& "File:" & vbCrLf & FileName
                End If
                retVals(k) = "0"
                k = k + 1
                
            End If
            
            mPos1 = mPos2
            If mPos1 = Len(defineLine) Then Exit Do
            
        Loop
    'samo jedna varijabla
    Else
    
    End If
End Sub
'provjera ispravnosti define linije
'Private Function checkDefine(ByVal defLn As String, ByRef errDescr As String) As Boolean
'    checkDefine = True
'    'ako ima više od jedne ; tada nevalja
'    If Mid$(defLn, Len(defLn), 1) <> LNG_ARGDELIMITER Then
'        checkDefine = False
'        errDescr = ERR_ExpEndOfLn
'        lastErrNum = ERR_ExpEndOfLnN
'    ElseIf getInstrCnt(defLn, "(") <> getInstrCnt(defLn, ")") Then
'        errDescr = ERR_ExpEndOfSt
'        lastErrNum = ERR_ExpEndOfStN
'        checkDefine = False
'    End If
'End Function
'///////////////////////////////////////////////////////////////////////
'               return
'///////////////////////////////////////////////////////////////////////
'Public Function checkReturn(ByVal retLn As String, ByRef errDescr As String) As Boolean
'    checkReturn = True
'    'nema ; na kraju
'    If Mid$(retLn, Len(retLn), 1) <> LNG_ARGDELIMITER Then
'        checkReturn = False
'        errDescr = ERR_ExpEndOfLn
'        lastErrNum = ERR_ExpEndOfLnN
'    '
'    ElseIf getInstrCnt(retLn, "(") <> getInstrCnt(retLn, ")") Then
'        errDescr = ERR_ExpEndOfSt
'        lastErrNum = ERR_ExpEndOfStN
'        checkReturn = False
'    End If
'End Function


'///////////////////////////////////////////////////////////////////////
'               math line
'///////////////////////////////////////////////////////////////////////
Public Sub retMathLine(ByVal mthLn As String, ByRef varNm As String, ByRef mathCode As String)
    
    Dim pos1 As String, errDesc As String
    Dim mPos1 As Integer
    If Mid$(mthLn, Len(mthLn), 1) = LNG_ARGDELIMITER Then
        mthLn = Mid$(mthLn, 1, Len(mthLn) - 1)
    End If
    
    'brisanje praznina
    If InStrB(1, mthLn, " ") <> 0 Then
    
        mthLn = Replace$(mthLn, " ", vbNullString)
        
    End If
    
    'traženje znaka =
    mPos1 = InStr(1, mthLn, "=")
    varNm = Mid$(mthLn, 1, mPos1 - 1)
    mathCode = Mid$(mthLn, mPos1 + 1)
'    If checkVarName(varNm) = False Then
'        Err.Raise -11, "clsCalculator", "Invalid variable '" & varNm & "'" & vbCrLf & vbCrLf & "File:" & vbCrLf & fileName
'    End If
End Sub

'Private Function checkMathLine(ByVal mathLine As String, ByRef errDescr As String) As Boolean
'    Dim pos1 As Integer, varNm As String, mathCode As String
'    checkMathLine = True
'    pos1 = Instr$(1, mathLine, OP_IS)
'    varNm = Mid$(mathLine, 1, pos1 - 1)
'    mathCode = Trim$(Mid$(mathLine, pos1 + 1))
'
'    If getInstrCnt(mathLine, OP_IS) > 1 Then
'        checkMathLine = False
'    ElseIf pos1 < 2 Then
'        checkMathLine = False
'        errDescr = ERR_ExpectedVarNm
'        lastErrNum = ERR_ExpectedVarNmN
'    ElseIf Mid$(mathLine, Len(mathLine), 1) <> LNG_ARGDELIMITER Then
'        checkMathLine = False
'        errDescr = ERR_ExpEndOfLn
'        lastErrNum = ERR_ExpEndOfLnN
'    ElseIf varNm = "$" Then
'        checkMathLine = False
'        errDescr = ERR_InvalidVarNm
'        lastErrNum = ERR_InvalidVarNmN
'    ElseIf mathCode = ";" Or Len(Trim$(mathCode)) = 0 Then
'        checkMathLine = False
'        errDescr = ERR_ExpectedExpres
'        lastErrNum = ERR_ExpectedExpresN
'    End If
'End Function
'///////////////////////////////////////////////////////////////////////
'               if
'///////////////////////////////////////////////////////////////////////
Public Sub getIF(ByVal ifLine As String, ByRef retIF As cmdIF)

    Dim mPos1 As Integer, mPos2 As Integer
    mPos1 = InStr(1, ifLine, "(")
    mPos2 = InStrRev(ifLine, ")")
    retIF.ifCondition = Mid$(ifLine, mPos1 + 1, mPos2 - mPos1 - 1)
    
End Sub
'
'Private Function checkIF(ByVal ifLine As String, ByRef errDescr As String) As Boolean
'    checkIF = True
'    If getInstrCnt(ifLine, "(") <> getInstrCnt(ifLine, ")") Then
'        errDescr = ERR_ExpEndOfSt
'        lastErrNum = ERR_ExpEndOfStN
'        checkIF = False
'    End If
'End Function
'///////////////////////////////////////////////////////////////////////
'               while
'///////////////////////////////////////////////////////////////////////
Public Sub getWhile(ByVal whileLine As String, ByRef mWhile As cmdWhile)

    Dim mPos1 As Integer, mPos2 As Integer
    mPos1 = InStr(1, whileLine, "(")
    mPos2 = InStrRev(whileLine, ")")
    
    mWhile.whileCondition = Mid$(whileLine, mPos1 + 1, mPos2 - mPos1 - 1)
    
End Sub
'Private Function checkWhile(ByVal whileLine As String, ByRef errDescr As String) As Boolean
'    checkWhile = True
'    If getInstrCnt(whileLine, "(") <> getInstrCnt(whileLine, ")") Then
'        errDescr = ERR_ExpEndOfSt
'        lastErrNum = ERR_ExpEndOfStN
'        checkWhile = False
'    End If
'End Function

'///////////////////////////////////////////////////////////////////////
'               vars
'///////////////////////////////////////////////////////////////////////
'
Public Function checkVarName(ByVal varName As String) As Boolean

    checkVarName = True
    If InStrB(1, varName, " ") > 0 Then
        checkVarName = False
        
    ElseIf InStrB(1, varName, "&") > 0 Then
        checkVarName = False
        
    ElseIf InStrB(1, varName, OP_IS) > 0 Then
        checkVarName = False
        
    ElseIf InStrB(1, varName, OP_PLUS) > 0 Then
        checkVarName = False
        
    ElseIf InStrB(1, varName, OP_MINUS) > 0 Then
        checkVarName = False
        
    ElseIf InStrB(1, varName, OP_MUL) > 0 Then
        checkVarName = False
        
    ElseIf InStrB(1, varName, OP_DIV) > 0 Then
        checkVarName = False
        
    ElseIf InStrB(1, varName, OP_POT) > 0 Then
        checkVarName = False
        
    ElseIf InStrB(1, varName, "%") > 0 Then
        checkVarName = False
        
    ElseIf InStrB(1, varName, OP_LESSTHAN) > 0 Then
        checkVarName = False
        
    ElseIf InStrB(1, varName, OP_GREATERTHAN) > 0 Then
        checkVarName = False
        
    ElseIf InStrB(1, varName, "!") > 0 Then
        checkVarName = False
        
    ElseIf InStrB(1, varName, "\") > 0 Then
        checkVarName = False
        
    ElseIf InStrB(1, varName, "(") > 0 Then
        checkVarName = False
        
    ElseIf InStrB(1, varName, ")") > 0 Then
        checkVarName = False
        
    ElseIf InStrB(1, varName, "|") > 0 Then
        checkVarName = False
        
    End If
    
End Function
'
'pretprovjera koda
'   ako vrati true onda je sve ok
'Public Function checkCode(ByVal strLn As String, ByRef ifLvl As Integer, ByRef whileLvl As Integer, ByRef lvlStog As Collection) As Boolean
'    'ako ima zabranjenih simbola
'    If haveInvalidChars(strLn) = True Then
'        lastErr = ERR_InvalidChars
'        lastErrNum = ERR_InvalidCharsN
'        checkCode = False
'        Exit Function
'    End If
'    '
'    checkCode = True
'    Dim errDesc As String
'    If StrComp(Mid$(strLn, 1, Len("return")), "return", vbTextCompare) = 0 Then
'        If checkReturn(strLn, errDesc) = False Then
'            lastErr = errDesc  '& strLn & vbCrLf & vbCrLf & "Function:" & vbCrLf & tmpFile
'            checkCode = False
'        End If
'    'define
'    ElseIf StrComp(Mid$(strLn, 1, LEN_DEFINE), "define", vbTextCompare) = 0 Then
'        If checkDefine(strLn, errDesc) = False Then
'            lastErr = errDesc '& strLn & vbCrLf & vbCrLf & "Function:" & vbCrLf & tmpFile
'            checkCode = False
'        End If
'    'if
'    ElseIf StrComp(Mid$(strLn, 1, Len("if")), "if", vbTextCompare) = 0 Then
'        ifLvl = ifLvl + 1
'        lvlStog.Add "IF"
'        If checkIF(strLn, errDesc) = False Then
'            lastErr = errDesc '& strLn & vbCrLf & vbCrLf & "Function:" & vbCrLf & tmpFile
'            checkCode = False
'        End If
'       ' MsgBox "IF STAVLJA"
'    'else if
'    ElseIf StrComp(Mid$(strLn, 1, Len("elseif")), "elseif", vbTextCompare) = 0 Or StrComp(Mid$(strLn, 1, Len("else if")), "else if", vbTextCompare) = 0 Then
'        If ifLvl = 0 Then
'            lastErr = "Elseif before block if"
'            checkCode = False
'        ElseIf lvlStog.Count = 0 Then
'            lastErr = "Elseif before block if"
'            checkCode = False
'        ElseIf lvlStog.Item(lvlStog.Count) <> "IF" Then
'            lastErr = "Elseif before block if"
'            checkCode = False
'        ElseIf checkIF(strLn, errDesc) = False Then
'            lastErr = errDesc '& strLn & vbCrLf & vbCrLf & "Function:" & vbCrLf & tmpFile
'            checkCode = False
'        End If
'        'lvlStog.Add "ELSEIF"
'    'else
'    ElseIf StrComp(Mid$(strLn, 1, Len("else")), "else", vbTextCompare) = 0 Then
'        If StrComp(strLn, "else", vbTextCompare) <> 0 Then
'            lastErr = "Unknow command:" & vbCrLf & strLn
'            checkCode = False
'        ElseIf ifLvl = 0 Then
'            lastErr = "Else before block if"
'            checkCode = False
'        ElseIf lvlStog.Count = 0 Then
'            lastErr = "Else before block if"
'            checkCode = False
'        ElseIf lvlStog.Item(lvlStog.Count) <> "IF" Then
'            lastErr = "Else before block if"
'            checkCode = False
'        End If
'    'end if
'    ElseIf StrComp(Mid$(strLn, 1, Len("end if")), "end if", vbTextCompare) = 0 Or StrComp(Mid$(strLn, 1, Len("endif")), "endif", vbTextCompare) = 0 Then
'        If StrComp(strLn, "end if", vbTextCompare) <> 0 And StrComp(strLn, "endif", vbTextCompare) <> 0 Then
'            lastErr = "Unknow command:" & vbCrLf & strLn
'            checkCode = False
'        ElseIf ifLvl = 0 Then
'            lastErr = "End if before block if"
'            checkCode = False
'        ElseIf lvlStog.Count = 0 Then
'            lastErr = "Else before block if"
'            checkCode = False
'        ElseIf lvlStog.Item(lvlStog.Count) <> "IF" Then
'            lastErr = "End if before block if"
'            checkCode = False
'        Else
'            lvlStog.Remove lvlStog.Count
'            'MsgBox "BRISE IF"
'        End If
'        ifLvl = ifLvl - 1
'    'while
'    ElseIf StrComp(Mid$(strLn, 1, Len("while")), "while", vbTextCompare) = 0 Then
'        whileLvl = whileLvl + 1
'        lvlStog.Add "WHILE"
'        'MsgBox "STAVLJA WH"
'        If checkWhile(strLn, errDesc) = False Then
'            lastErr = errDesc '& strLn & vbCrLf & vbCrLf & "Function:" & vbCrLf & tmpFile
'            checkCode = False
'        End If
'    'loop
'    ElseIf StrComp(Mid$(strLn, 1, Len("loop")), "loop", vbTextCompare) = 0 Then
'        If StrComp(strLn, "loop", vbTextCompare) <> 0 And StrComp(strLn, "loop;", vbTextCompare) <> 0 Then
'            lastErr = "Unknow command:" & vbCrLf & strLn
'            checkCode = False
'        ElseIf whileLvl = 0 Then
'            lastErr = "Expected while before end loop"
'            checkCode = False
'        ElseIf lvlStog.Count = 0 Then
'            lastErr = "loop before block while"
'            checkCode = False
'        ElseIf lvlStog.Item(lvlStog.Count) <> "WHILE" Then
'            lastErr = "loop before block while"
'            checkCode = False
'        Else
'            lvlStog.Remove lvlStog.Count
'         '   MsgBox "BRISE WHILE"
'        End If
'        whileLvl = whileLvl - 1
'    'break
'    ElseIf StrComp(Mid$(strLn, 1, Len("break")), "break", vbTextCompare) = 0 Then
'       If StrComp(strLn, "break", vbTextCompare) <> 0 And StrComp(strLn, "break;", vbTextCompare) <> 0 Then
'            lastErr = "Unknow command:" & vbCrLf & strLn
'            checkCode = False
'        ElseIf whileLvl = 0 Then
'            lastErr = "Expected while before end break"
'            checkCode = False
'        ElseIf lvlStog.Count = 0 Then
'            lastErr = "break before block while"
'            checkCode = False
'        ElseIf StrComp(lvlStog.Item(lvlStog.Count), "while", vbTextCompare) <> 0 Then
'            lastErr = "break before block while"
'            checkCode = False
'        End If
'
'    'matematièke operacije
'    ElseIf Instr$(1, strLn, OP_IS) > 0 Then
'        If checkMathLine(strLn, errDesc) = False Then
'            lastErr = errDesc '& strLn & vbCrLf & vbCrLf & "Function:" & vbCrLf & tmpFile
'            checkCode = False
'        End If
'    Else
'        lastErr = "Sintax error or unknow command:" & vbCrLf & strLn
'        checkCode = False
'    End If
'End Function

'ako ima neke znakove koje ne smije imati
Public Function haveInvalidChars(ByVal mLine As String) As Boolean
    
    haveInvalidChars = False
    
    If InStrB(1, mLine, "\") Then
    
        haveInvalidChars = True
        
    ElseIf InStrB(1, mLine, ":") Then
    
        haveInvalidChars = True
        
    ElseIf InStrB(1, mLine, "?") Then
    
        haveInvalidChars = True
        
    ElseIf InStrB(1, mLine, "'") Then
    
        haveInvalidChars = True
        
    ElseIf InStrB(1, mLine, "()") Then
    
        haveInvalidChars = True
        
    ElseIf InStrB(1, mLine, "=<") Then
    
        haveInvalidChars = True
        
    ElseIf InStrB(1, mLine, "=>") Then
    
        haveInvalidChars = True
        
    End If
    
End Function


