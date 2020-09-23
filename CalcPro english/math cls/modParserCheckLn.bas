Attribute VB_Name = "modParserCheckLn"
Option Explicit

Public Function checkLineSintax(ByVal strLn As String, ByRef ifLvl As Integer, ByRef whileLvl As Integer, ByVal lvlStog As Collection) As Boolean
    
    Dim mPos1 As Integer, mPos2 As Integer
    Dim mStart1 As Integer, mEnd1 As Integer
    Dim tmpStr As String, tmpStr2 As String
    checkLineSintax = True
    '
    'ako imamo razlicit broj ( i )
    '   onda je greška
    If getInstrCnt(strLn, "(") <> getInstrCnt(strLn, ")") Then
    
        lastErr = ERR_ExpEndOfSt & ":" & vbCrLf & strLn
        lastErrNum = ERR_ExpEndOfStN
        checkLineSintax = False
        Exit Function
        
    End If
    '------------------------------------------------------------------------------
    '*************************  FUNCTION ******************************************
    '------------------------------------------------------------------------------
    If StrComp(Mid$(strLn, 1, LEN_FUNCTION), "function", vbTextCompare) = 0 Then
        mPos1 = InStr(1, strLn, " ")
        mPos2 = InStrRev(strLn, ")")
        
        If mPos1 = 0 Or mPos2 = 0 Then
            lastErr = ERR_UnknownCmd & ": " & vbCrLf & strLn
            lastErrNum = ERR_UnknownCmdN
            checkLineSintax = False
            Exit Function
        End If
        'provjeravamo da li se išta nalazi iza function
        '   što nebi trebalo
        tmpStr = Trim$(Mid$(strLn, 1, mPos1 - 1))
        'MsgBox tmpStr
        If StrComp(tmpStr, "function", vbTextCompare) <> 0 Then
            lastErr = ERR_ExpEndOfSt & ":" & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfStN
            checkLineSintax = False
            Exit Function
        End If
        mPos1 = InStr(1, strLn, "(")
        'provjeravamo da li se išta nalazi iza function(;;;)
        tmpStr2 = Trim$(Mid$(strLn, mPos2 + 1))
        'MsgBox tmpStr2
        If Len(tmpStr2) > 0 Then
            lastErr = ERR_ExpEndOfSt & ":" & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfStN
            checkLineSintax = False
            Exit Function
        End If
        
        'provjerimo argumente, tj. ono
        '   što se nalazi u zagradama
        tmpStr2 = Trim$(Mid$(strLn, mPos1 + 1, mPos2 - mPos1 - 1))
        If Len(tmpStr2) = 0 Then
            lastErr = ERR_InvalidArg & ":" & vbCrLf & strLn
            lastErrNum = ERR_InvalidArgN
            checkLineSintax = False
            Exit Function
        End If
        
        mStart1 = InStr(1, tmpStr2, ";")
        'ako ima više argumenata
        If mStart1 > 0 Then
            mStart1 = 0
            Do 'While mStart1 > 0
                mEnd1 = InStr(mStart1 + 1, tmpStr2, ";")
                If mEnd1 = 0 Then
                    mEnd1 = Len(tmpStr2) + 1
                End If
                tmpStr = Trim$(Mid$(tmpStr2, mStart1 + 1, mEnd1 - 1))
                '
                'provjera argumenta
                'ako nema $ ispred naziva, greška
                If left$(tmpStr, 1) <> "$" Then
                    lastErr = ERR_InvalidArg & ": " & tmpStr & vbCrLf & "Line:" & strLn
                    lastErrNum = ERR_InvalidArgN
                    checkLineSintax = False
                'ako sadrži zabranjene znakove
                ElseIf haveInvalidChars(tmpStr) = True Then
                    lastErr = ERR_InvalidArg & ": " & tmpStr & vbCrLf & "Line:" & strLn
                    lastErrNum = ERR_InvalidArgN
                    checkLineSintax = False
                'ili ako sadrži zarez (isto zabranjen vamo)
                ElseIf InStrB(1, tmpStr, ",") <> 0 Then
                    lastErr = ERR_InvalidArg & ": " & tmpStr & vbCrLf & "Line:" & strLn
                    lastErrNum = ERR_InvalidArgN
                    checkLineSintax = False
                End If
                '
                mStart1 = mEnd1
                If mEnd1 >= Len(tmpStr2) Then
                    Exit Do
                End If
            Loop
        'ako je samo jedan ili nijedan
        Else
            If Len(tmpStr2) > 0 Then
                'provjera argumenta
                'ako nema $ ispred naziva, greška
                If left$(tmpStr2, 1) <> "$" Then
                    lastErr = ERR_InvalidArg & ": " & tmpStr2 & vbCrLf & "Line:" & strLn
                    lastErrNum = ERR_InvalidArgN
                    checkLineSintax = False
                'ako sadrži zabranjene znakove
                ElseIf haveInvalidChars(tmpStr2) = True Then
                    lastErr = ERR_InvalidArg & ": " & tmpStr2 & vbCrLf & "Line:" & strLn
                    lastErrNum = ERR_InvalidArgN
                    checkLineSintax = False
                'ili ako sadrži zarez (isto zabranjen vamo)
                ElseIf InStrB(1, tmpStr2, ",") <> 0 Then
                    lastErr = ERR_InvalidArg & ": " & tmpStr2 & vbCrLf & "Line:" & strLn
                    lastErrNum = ERR_InvalidArgN
                    checkLineSintax = False
                End If
            End If
        End If
    '------------------------------------------------------------------------------
    '*************************  define ******************************************
    '------------------------------------------------------------------------------
    ElseIf StrComp(Mid$(strLn, 1, LEN_DEFINE), "define", vbTextCompare) = 0 Then
        mPos1 = InStr(1, strLn, " ")
        mPos2 = InStrRev(strLn, ";")
        'ako ima više od jednog ;
        If getInstrCnt(strLn, ";") > 1 Then
            lastErr = ERR_ExpEndOfSt & ": " & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfStN
            checkLineSintax = False
            Exit Function
        ElseIf mPos2 = 0 Then
            lastErr = ERR_ExpEndOfLn & ": " & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfLnN
            checkLineSintax = False
            Exit Function
        ElseIf mPos1 = 0 Then
            lastErr = ERR_UnknownCmd & ": " & vbCrLf & strLn
            lastErrNum = ERR_UnknownCmdN
            checkLineSintax = False
            Exit Function
        End If
        '
        'ispravnost naredbe
        tmpStr = Trim$(Mid$(strLn, 1, mPos1 - 1))
        If StrComp(tmpStr, "define", vbTextCompare) <> 0 Then
            lastErr = ERR_UnknownCmd & ": " & vbCrLf & strLn
            lastErrNum = ERR_UnknownCmdN
            checkLineSintax = False
            Exit Function
        End If
        
        'provjeravamo da li se išta nalazi iza define $,,,,;
        tmpStr2 = Trim$(Mid$(strLn, mPos2 + 1))
        'MsgBox tmpStr2
        If Len(tmpStr2) > 0 Then
            lastErr = ERR_ExpEndOfSt & ":" & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfStN
            checkLineSintax = False
            Exit Function
        End If
        
        'provjerimo varijable
        tmpStr2 = Trim$(Mid$(strLn, mPos1 + 1, mPos2 - mPos1 - 1))
        mStart1 = InStr(1, tmpStr2, ",")
        'ako ima više argumenata
        If mStart1 > 0 Then
            mStart1 = 0 'mPos1
            Do 'While mStart1 > 0
                mEnd1 = InStr(mStart1 + 1, tmpStr2, ",")
                If mEnd1 = 0 Then
                    mEnd1 = Len(tmpStr2) + 1
                End If
                tmpStr = Trim$(Mid$(tmpStr2, mStart1 + 1, mEnd1 - 1))
                'MsgBox mStart1 & vbCrLf & mEnd1 & vbCrLf & tmpStr
                '
                'provjera argumenta
                'ako nema $ ispred naziva, greška
                If left$(tmpStr, 1) <> "$" Then
                    lastErr = ERR_InvalidArg & ": " & tmpStr & vbCrLf & "Line:" & strLn
                    lastErrNum = ERR_InvalidArgN
                    checkLineSintax = False
                'ako sadrži zabranjene znakove
                ElseIf haveInvalidChars(tmpStr) = True Then
                    lastErr = ERR_InvalidArg & ": " & tmpStr & vbCrLf & "Line:" & strLn
                    lastErrNum = ERR_InvalidArgN
                    checkLineSintax = False
                'ili ako sadrži zarez (isto zabranjen vamo)
                ElseIf InStrB(1, tmpStr, ";") <> 0 Then
                    lastErr = ERR_InvalidArg & ": " & tmpStr & vbCrLf & "Line:" & strLn
                    lastErrNum = ERR_InvalidArgN
                    checkLineSintax = False
                End If
                '
                mStart1 = mEnd1
                If mEnd1 >= Len(tmpStr2) Then
                    Exit Do
                End If
            Loop
        'ako je samo jedan ili nijedan
        Else
            'provjera argumenta
            'ako nema $ ispred naziva, greška
            If left$(tmpStr2, 1) <> "$" Then
                lastErr = ERR_InvalidArg & ": " & tmpStr2 & vbCrLf & "Line:" & strLn
                lastErrNum = ERR_InvalidArgN
                checkLineSintax = False
            'ako sadrži zabranjene znakove
            ElseIf haveInvalidChars(tmpStr2) = True Then
                lastErr = ERR_InvalidArg & ": " & tmpStr2 & vbCrLf & "Line:" & strLn
                lastErrNum = ERR_InvalidArgN
                checkLineSintax = False
            'ili ako sadrži zarez (isto zabranjen vamo)
            ElseIf InStrB(1, tmpStr2, ";") <> 0 Then
                lastErr = ERR_InvalidArg & ": " & tmpStr2 & vbCrLf & "Line:" & strLn
                lastErrNum = ERR_InvalidArgN
                checkLineSintax = False
            End If
        End If
    '------------------------------------------------------------------------------
    '*************************  if ******************************************
    '------------------------------------------------------------------------------
    ElseIf StrComp(Mid$(strLn, 1, 2), "if", vbTextCompare) = 0 Then
        ifLvl = ifLvl + 1
        lvlStog.Add "IF"
        '
        mPos1 = InStr(1, strLn, "(")
        mPos2 = InStrRev(strLn, ")")
        If mPos2 = 0 Then
            lastErr = ERR_UnknownCmd & ": " & vbCrLf & strLn
            lastErrNum = ERR_UnknownCmdN
            checkLineSintax = False
            Exit Function
        ElseIf mPos1 = 0 Then
            lastErr = ERR_UnknownCmd & ": " & vbCrLf & strLn
            lastErrNum = ERR_UnknownCmdN
            checkLineSintax = False
            Exit Function
        End If
        '
        'ispravnost naredbe
        tmpStr = Trim$(Mid$(strLn, 1, mPos1 - 1))
        If StrComp(tmpStr, "if", vbTextCompare) <> 0 Then
            lastErr = ERR_UnknownCmd & ": " & vbCrLf & strLn
            lastErrNum = ERR_UnknownCmdN
            checkLineSintax = False
            Exit Function
        End If
        
        'provjeravamo da li se išta nalazi iza if(any)
        tmpStr2 = Trim$(Mid$(strLn, mPos2 + 1))
        'MsgBox tmpStr2
        If Len(tmpStr2) > 0 Then
            lastErr = ERR_ExpEndOfSt & ":" & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfStN
            checkLineSintax = False
            Exit Function
        End If
        
        'provjerimo uvjet
        tmpStr2 = Trim$(Mid$(strLn, mPos1 + 1, mPos2 - mPos1 - 1))
        If haveInvalidChars(tmpStr2) = True Then
            lastErr = ERR_InvalidChars & ":" & vbCrLf & strLn
            lastErrNum = ERR_InvalidCharsN
            checkLineSintax = False
            Exit Function
        'dead code
        ElseIf tmpStr2 = "0" Then
            lastErr = ERR_DeadCode & ":" & vbCrLf & strLn
            lastErrNum = ERR_DeadCodeN
            checkLineSintax = False
            Exit Function
        'nema izraza
        ElseIf Len(tmpStr2) = 0 Then
            lastErr = ERR_ExpectedExpres & ":" & vbCrLf & strLn
            lastErrNum = ERR_ExpectedExpresN
            checkLineSintax = False
            Exit Function
        End If
    '------------------------------------------------------------------------------
    '*************************  elseif ******************************************
    '------------------------------------------------------------------------------
    ElseIf StrComp(Mid$(strLn, 1, Len("elseif")), "elseif", vbTextCompare) = 0 Then
        'provjera li postoji if
        If ifLvl = 0 Then
            lastErr = ERR_ElseIfBefIf & ": " & vbCrLf & strLn
            lastErrNum = ERR_ElseIfBefIfN
            checkLineSintax = False
            Exit Function
        ElseIf lvlStog.Count = 0 Then
            lastErr = ERR_ElseIfBefIf & ": " & vbCrLf & strLn
            lastErrNum = ERR_ElseIfBefIfN
            checkLineSintax = False
            Exit Function
        ElseIf lvlStog.Item(lvlStog.Count) <> "IF" Then
            lastErr = ERR_ElseIfBefIf & ": " & vbCrLf & strLn
            lastErrNum = ERR_ElseIfBefIfN
            checkLineSintax = False
            Exit Function
        End If
        
        mPos1 = InStr(1, strLn, "(")
        mPos2 = InStrRev(strLn, ")")
        If mPos2 = 0 Then
            lastErr = ERR_UnknownCmd & ": " & vbCrLf & strLn
            lastErrNum = ERR_UnknownCmdN
            checkLineSintax = False
            Exit Function
        ElseIf mPos1 = 0 Then
            lastErr = ERR_UnknownCmd & ": " & vbCrLf & strLn
            lastErrNum = ERR_UnknownCmdN
            checkLineSintax = False
            Exit Function
        End If
        '
        'ispravnost naredbe
        tmpStr = Trim$(Mid$(strLn, 1, mPos1 - 1))
        If StrComp(tmpStr, "elseif", vbTextCompare) <> 0 Then
            lastErr = ERR_UnknownCmd & ": " & vbCrLf & strLn
            lastErrNum = ERR_UnknownCmdN
            checkLineSintax = False
            Exit Function
        End If
        
        'provjeravamo da li se išta nalazi iza if(any)
        tmpStr2 = Trim$(Mid$(strLn, mPos2 + 1))
        'MsgBox tmpStr2
        If Len(tmpStr2) > 0 Then
            lastErr = ERR_ExpEndOfSt & ":" & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfStN
            checkLineSintax = False
            Exit Function
        End If
        
        'provjerimo uvjet
        tmpStr2 = Trim$(Mid$(strLn, mPos1 + 1, mPos2 - mPos1 - 1))
        If haveInvalidChars(tmpStr2) = True Then
            lastErr = ERR_InvalidChars & ":" & vbCrLf & strLn
            lastErrNum = ERR_InvalidCharsN
            checkLineSintax = False
            Exit Function
        'dead code
        ElseIf tmpStr2 = "0" Then
            lastErr = ERR_DeadCode & ":" & vbCrLf & strLn
            lastErrNum = ERR_DeadCodeN
            checkLineSintax = False
            Exit Function
        'nema izraza
        ElseIf Len(tmpStr2) = 0 Then
            lastErr = ERR_ExpectedExpres & ":" & vbCrLf & strLn
            lastErrNum = ERR_ExpectedExpresN
            checkLineSintax = False
            Exit Function
        End If

    '------------------------------------------------------------------------------
    '*************************  else ******************************************
    '------------------------------------------------------------------------------
    ElseIf StrComp(Mid$(strLn, 1, Len("else")), "else", vbTextCompare) = 0 Then
        'provjera da li uopce postoji if
        If ifLvl = 0 Then
            lastErr = ERR_ElseBefIf & ":" & vbCrLf & strLn
            lastErrNum = ERR_ElseBefIfN
            checkLineSintax = False
            Exit Function
        ElseIf lvlStog.Count = 0 Then
            lastErr = ERR_ElseBefIf & ":" & vbCrLf & strLn
            lastErrNum = ERR_ElseBefIfN
            checkLineSintax = False
            Exit Function
        ElseIf lvlStog.Item(lvlStog.Count) <> "IF" Then
            lastErr = ERR_ElseBefIf & ":" & vbCrLf & strLn
            lastErrNum = ERR_ElseBefIfN
            checkLineSintax = False
            Exit Function
        End If
        '
        mPos2 = Len("ELSE")
        'provjeravamo da li se išta nalazi iza if(any)
        tmpStr2 = Trim$(Mid$(strLn, mPos2 + 1))
        If Len(tmpStr2) > 0 Then
            lastErr = ERR_ExpEndOfSt & ":" & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfStN
            checkLineSintax = False
            Exit Function
        End If
    '------------------------------------------------------------------------------
    '*************************  end if ******************************************
    '------------------------------------------------------------------------------
    ElseIf StrComp(Mid$(strLn, 1, Len("end if")), "end if", vbTextCompare) = 0 Then
        If ifLvl = 0 Then
            lastErr = ERR_ElseBefIf & ":" & vbCrLf & strLn
            lastErrNum = ERR_ElseBefIfN
            checkLineSintax = False
            Exit Function
        ElseIf lvlStog.Count = 0 Then
            lastErr = ERR_ElseBefIf & ":" & vbCrLf & strLn
            lastErrNum = ERR_ElseBefIfN
            checkLineSintax = False
            Exit Function
        ElseIf lvlStog.Item(lvlStog.Count) <> "IF" Then
            lastErr = ERR_ElseBefIf & ":" & vbCrLf & strLn
            lastErrNum = ERR_ElseBefIfN
            checkLineSintax = False
            Exit Function
        Else
            lvlStog.Remove lvlStog.Count
        End If
        ifLvl = ifLvl - 1
        
        mPos2 = Len("end if")
        'provjeravamo da li se išta nalazi iza if(any)
        tmpStr2 = Trim$(Mid$(strLn, mPos2 + 1))
        'MsgBox tmpStr2
        If Len(tmpStr2) > 0 Then
            lastErr = ERR_ExpEndOfSt & ":" & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfStN
            checkLineSintax = False
            Exit Function
        End If
    '------------------------------------------------------------------------------
    '*************************  while ******************************************
    '------------------------------------------------------------------------------
    ElseIf StrComp(Mid$(strLn, 1, Len("While")), "While", vbTextCompare) = 0 Then
        whileLvl = whileLvl + 1
        lvlStog.Add "WHILE"
        
        mPos1 = InStr(1, strLn, "(")
        mPos2 = InStrRev(strLn, ")")
        If mPos2 = 0 Then
            lastErr = ERR_UnknownCmd & ": " & vbCrLf & strLn
            lastErrNum = ERR_UnknownCmdN
            checkLineSintax = False
            Exit Function
        ElseIf mPos1 = 0 Then
            lastErr = ERR_UnknownCmd & ": " & vbCrLf & strLn
            lastErrNum = ERR_UnknownCmdN
            checkLineSintax = False
            Exit Function
        End If
        
        '
        'ispravnost naredbe
        tmpStr = Trim$(Mid$(strLn, 1, mPos1 - 1))
        If StrComp(tmpStr, "While", vbTextCompare) <> 0 Then
            lastErr = ERR_UnknownCmd & ": " & vbCrLf & strLn
            lastErrNum = ERR_UnknownCmdN
            checkLineSintax = False
            Exit Function
        End If
        
        'provjeravamo da li se išta nalazi iza if(any)
        tmpStr2 = Trim$(Mid$(strLn, mPos2 + 1))
        'MsgBox tmpStr2
        If Len(tmpStr2) > 0 Then
            lastErr = ERR_ExpEndOfSt & ":" & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfStN
            checkLineSintax = False
            Exit Function
        End If
        
        'provjerimo uvjet
        tmpStr2 = Trim$(Mid$(strLn, mPos1 + 1, mPos2 - mPos1 - 1))
        If haveInvalidChars(tmpStr2) = True Then
            lastErr = ERR_InvalidChars & ":" & vbCrLf & strLn
            lastErrNum = ERR_InvalidCharsN
            checkLineSintax = False
            Exit Function
        End If
        'nema uvjeta
        strLn = Replace$(strLn, " ", vbNullString)
        If StrComp(strLn, "while()", vbTextCompare) = 0 Then
            lastErr = ERR_ExpectedStatm & ":" & vbCrLf & strLn
            lastErrNum = ERR_ExpectedStatmN
            checkLineSintax = False
            Exit Function
        'dead code
        ElseIf tmpStr2 = "0" Then
            lastErr = ERR_DeadCode & ":" & vbCrLf & strLn
            lastErrNum = ERR_DeadCodeN
            checkLineSintax = False
            Exit Function
        'nema izraza
        ElseIf Len(tmpStr2) = 0 Then
            lastErr = ERR_ExpectedExpres & ":" & vbCrLf & strLn
            lastErrNum = ERR_ExpectedExpresN
            checkLineSintax = False
            Exit Function
        End If
    '------------------------------------------------------------------------------
    '*************************  loop ******************************************
    '------------------------------------------------------------------------------
    ElseIf StrComp(Mid$(strLn, 1, Len("loop")), "loop", vbTextCompare) = 0 Then
        If whileLvl = 0 Then
            'MsgBox "1"
            lastErr = ERR_LoopBefWhile & ":" & vbCrLf & strLn
            lastErrNum = ERR_LoopBefWhileN
            checkLineSintax = False
            Exit Function
        ElseIf lvlStog.Count = 0 Then
            'MsgBox "2"
            lastErr = ERR_LoopBefWhile & ":" & vbCrLf & strLn
            lastErrNum = ERR_LoopBefWhileN
            checkLineSintax = False
            Exit Function
        ElseIf lvlStog.Item(lvlStog.Count) <> "WHILE" Then
'            MsgBox lvlStog.Item(lvlStog.Count)
'            MsgBox "3"
            lastErr = ERR_LoopBefWhile & ":" & vbCrLf & strLn
            lastErrNum = ERR_LoopBefWhileN
            checkLineSintax = False
            Exit Function
        Else
            lvlStog.Remove lvlStog.Count
         '   MsgBox "BRISE WHILE"
        End If
        whileLvl = whileLvl - 1
        
        mPos2 = Len("loop")
        'provjeravamo da li se išta nalazi iza if(any)
        tmpStr2 = Trim$(Mid$(strLn, mPos2 + 1))
        If Len(tmpStr2) > 0 Then
            lastErr = ERR_ExpEndOfSt & ":" & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfStN
            checkLineSintax = False
            Exit Function
        End If
    '------------------------------------------------------------------------------
    '************************* break ******************************************
    '------------------------------------------------------------------------------
    ElseIf StrComp(Mid$(strLn, 1, Len("break")), "break", vbTextCompare) = 0 Then
        mPos1 = InStrRev(strLn, ";")
        mPos2 = Len("break")
        'prvo provjerimo da li postoji
        '   while, ako ne postoji, greška
        If whileLvl = 0 Then
            lastErr = ERR_BreakWthWhile & ":" & vbCrLf & strLn
            lastErrNum = ERR_BreakWthWhileN
            checkLineSintax = False
            Exit Function
        ElseIf lvlStog.Count = 0 Then
            lastErr = ERR_BreakWthWhile & ":" & vbCrLf & strLn
            lastErrNum = ERR_BreakWthWhileN
            checkLineSintax = False
            Exit Function
        End If
        
        'ako ne ; na kraju
        If mPos1 = 0 Then
            lastErr = ERR_ExpEndOfLn & ": " & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfLnN
            checkLineSintax = False
            Exit Function
        End If
        
        'ako je naredba oblika  break any;
        tmpStr = Trim$(Mid$(strLn, 1, mPos2))
        If StrComp(tmpStr, "break", vbTextCompare) <> 0 Then
            lastErr = ERR_UnknownCmd & ": " & vbCrLf & strLn
            lastErrNum = ERR_UnknownCmdN
            checkLineSintax = False
            Exit Function
        End If
        'provjeravamo da li se išta nalazi iza break
        tmpStr2 = Trim$(Mid$(strLn, mPos1 + 1))
        If Len(tmpStr2) > 0 Then
            lastErr = ERR_ExpEndOfSt & ":" & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfStN
            checkLineSintax = False
            Exit Function
        End If
    '------------------------------------------------------------------------------
    '*************************  math line *****************************************
    '------------------------------------------------------------------------------
    ElseIf InStrB(1, strLn, "=") <> 0 Then
        mPos1 = InStr(1, strLn, "=")
        mPos2 = InStrRev(strLn, ";")
        
        If mPos2 = 0 Then
            lastErr = ERR_ExpEndOfLn & ": " & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfLnN
            checkLineSintax = False
            Exit Function
        End If
        'provjera varijable
        tmpStr = Trim$(left$(strLn, mPos1 - 1))
        If left$(tmpStr, 1) <> "$" Then
            lastErr = ERR_InvalidVarNm & ":" & vbCrLf & strLn
            lastErrNum = ERR_InvalidVarNmN
            checkLineSintax = False
            Exit Function
        ElseIf tmpStr = "$" Then
            lastErr = ERR_InvalidVarNm & ":" & vbCrLf & strLn
            lastErrNum = ERR_InvalidVarNmN
            checkLineSintax = False
            Exit Function
        End If
        'provjera linije koda
        tmpStr2 = Mid$(strLn, mPos1 + 1, mPos2 - mPos1 - 1)
        If haveInvalidChars(tmpStr2) = True Then
            lastErr = ERR_InvalidChars & ":" & vbCrLf & strLn
            lastErrNum = ERR_InvalidCharsN
            checkLineSintax = False
            Exit Function
        End If
        '
        'provjeravamo da li se išta nalazi iza $var=any;
        tmpStr2 = Trim$(Mid$(strLn, mPos2 + 1))
        If Len(tmpStr2) > 0 Then
            lastErr = ERR_ExpEndOfSt & ":" & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfStN
            checkLineSintax = False
            Exit Function
        End If
    '------------------------------------------------------------------------------
    '*************************  return *****************************************
    '------------------------------------------------------------------------------
    ElseIf StrComp(Mid$(strLn, 1, Len("return")), "return", vbTextCompare) = 0 Then
        mPos1 = InStr(1, strLn, " ")
        mPos2 = InStrRev(strLn, ";")
        If mPos1 = 0 Or mPos1 > mPos2 Then mPos1 = mPos2
        
        If mPos2 = 0 Then
            lastErr = ERR_ExpEndOfLn & ": " & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfLnN
            checkLineSintax = False
            Exit Function
'        ElseIf mPos1 = 0 Then
'            lastErr = ERR_UnknownCmd & ":" & vbCrLf & strLn
'            lastErrNum = ERR_UnknownCmdN
'            checkLineSintax = False
'            Exit Function
        End If
        'provjera varijable
        tmpStr = Trim$(left$(strLn, mPos1 - 1))
        If StrComp(tmpStr, "return", vbTextCompare) <> 0 Then
            lastErr = ERR_ExpEndOfSt & ": " & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfStN
            'MsgBox tmpStr
            checkLineSintax = False
            Exit Function
        End If
        
        tmpStr = Trim$(Mid$(strLn, mPos2 + 1))
        'MsgBox tmpStr
        If Len(tmpStr) > 0 Then
            lastErr = ERR_ExpEndOfSt & ": " & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfStN
            checkLineSintax = False
            Exit Function
        End If
        'provjerimo ima li što za vracanje, ako ima,
        '   provjerimo to što se vraca
        If mPos1 <> mPos2 Then
            'provjera linije koda
            tmpStr2 = Mid$(strLn, mPos1 + 1, mPos2 - mPos1 - 1)
            If haveInvalidChars(tmpStr2) = True Then
                lastErr = ERR_InvalidChars & ":" & vbCrLf & strLn
                lastErrNum = ERR_InvalidCharsN
                checkLineSintax = False
                Exit Function
            End If
        End If
    'end function
    ElseIf StrComp(Mid$(strLn, 1, Len("end function")), "end function", vbTextCompare) = 0 Then
        mPos2 = Len("end function")
        'provjeravamo da li se išta nalazi iza if(any)
        tmpStr2 = Trim$(Mid$(strLn, mPos2 + 1))
        'MsgBox tmpStr2
        If Len(tmpStr2) > 0 Then
            lastErr = ERR_ExpEndOfSt & ":" & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfStN
            checkLineSintax = False
            Exit Function
        End If
    '------------------------------------------------------------------------------
    '**************************  message ******************************************
    '------------------------------------------------------------------------------
    ElseIf StrComp(Mid$(strLn, 1, Len("message")), "message", vbTextCompare) = 0 Then
        mPos1 = InStr(1, strLn, "(")
        mPos2 = InStrRev(strLn, ";")
        If mPos1 = 0 Or mPos1 > mPos2 Then mPos1 = mPos2
        
        If mPos2 = 0 Then
            lastErr = ERR_ExpEndOfLn & ": " & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfLnN
            checkLineSintax = False
            Exit Function
        End If

        tmpStr = Trim$(Mid$(strLn, mPos2 + 1))
        If Len(tmpStr) > 0 Then
            lastErr = ERR_ExpEndOfSt & ": " & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfStN
            checkLineSintax = False
            Exit Function
        End If
        
        If getInstrCnt(strLn, "(") <> getInstrCnt(strLn, ")") Then
            lastErr = ERR_ExpEndOfSt & ": " & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfStN
            checkLineSintax = False
            Exit Function
        End If
    '------------------------------------------------------------------------------
    '**************************  set err ******************************************
    '------------------------------------------------------------------------------
    ElseIf StrComp(Mid$(strLn, 1, Len("setError")), "setError", vbTextCompare) = 0 Then
        mPos1 = InStr(1, strLn, "(")
        mPos2 = InStrRev(strLn, ";")
        If mPos1 = 0 Or mPos1 > mPos2 Then mPos1 = mPos2
        
        If mPos2 = 0 Then
            lastErr = ERR_ExpEndOfLn & ": " & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfLnN
            checkLineSintax = False
            Exit Function
        End If

        tmpStr = Trim$(Mid$(strLn, mPos2 + 1))
        If Len(tmpStr) > 0 Then
            lastErr = ERR_ExpEndOfSt & ": " & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfStN
            checkLineSintax = False
            Exit Function
        End If
        
        If getInstrCnt(strLn, "(") <> getInstrCnt(strLn, ")") Then
            lastErr = ERR_ExpEndOfSt & ": " & vbCrLf & strLn
            lastErrNum = ERR_ExpEndOfStN
            checkLineSintax = False
            Exit Function
        End If
    Else
        lastErr = ERR_UnknownCmd & ":" & vbCrLf & strLn
        lastErrNum = ERR_UnknownCmdN
        checkLineSintax = False
        Exit Function
    End If
End Function
