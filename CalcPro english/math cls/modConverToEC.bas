Attribute VB_Name = "modConverToEC"
Option Explicit

Public Sub convertToExecutable(ByRef codeHolder As Collection, ByRef retExeCode As Collection)
    
    Dim i As Integer, j As Integer, cnt1 As Integer
    Dim whileCnt As Integer
    Dim prevIf As Integer
    Dim strLn As String, tmpVarNms() As String, tmpVarVals() As String
    Dim tmpStr As String
    Dim mIf As cmdIF
    
    Dim mWhile As cmdWhile
    'conditions
    Dim whileStog As New Collection, whileCond As New Collection
    

     whileCnt = 0
    'ako se u execute kolekciji nešto nalazi
    '   brišemo to
    Set retExeCode = New Collection
    '
    'idemo rješavati liniju po liniju
    'For Each vItm In codeHolder
    cnt1 = codeHolder.Count
    
    For i = 1 To cnt1
    
        strLn = codeHolder.Item(i)
        'function
        If InStrB(1, strLn, "function", vbTextCompare) <> 0 Then ''StrComp(Mid$(strLn, 1, Len("break")), "break", vbTextCompare) = 0 Then
            
            retExeCode.Add Trim$(strLn)
            
        ElseIf InStrB(1, strLn, "return", vbTextCompare) <> 0 Then
            
            'stavljamo liniju: ret:value
            tmpStr = Trim$(strLn)
            If right$(tmpStr, 1) = ";" Then
                tmpStr = Mid$(tmpStr, 1, Len(tmpStr) - 1)
            End If
            retExeCode.Add Replace$((Replace$(tmpStr, "return", "ret:", , , vbTextCompare)), " ", vbNullString)
        
        'define
        ElseIf InStrB(1, strLn, "define", vbTextCompare) <> 0 Then
            
            'svaka varijabla u novi red
            getDefineVars strLn, tmpVarNms, tmpVarVals
            For j = 0 To UBound(tmpVarNms)
                retExeCode.Add "def:" & Trim$(tmpVarNms(j)) & "=" & Trim$(tmpVarVals(j))
            Next j
        
        'else if
        ElseIf InStrB(1, strLn, "elseif", vbTextCompare) <> 0 Then
            
            'prvo stavimo kod koji preskace sve do end if
            '   ukoliko je blok prije izvršen
            retExeCode.Add "JPF:" & getNextIfElseIf(codeHolder, i, True)
            'izvadimo uvjet
            getIF strLn, mIf
            'napravimo kod koji ce naci
            '   vrijednost uvjeta
            retExeCode.Add "$tmp_IfSYSVAR_i=" & Replace$(mIf.ifCondition, " ", vbNullString)
            'stavimo kod koji ce preskociti ukoliko je end if
            retExeCode.Add "JP:$tmp_IfSYSVAR_i==0:" & getNextIfElseIf(codeHolder, i)
            '
            '
'            'izmjena na prvobitnom if ili elseif prije ovog
'            prevIf = condStog.Item(condStog.Count)
'            'dodamo novi JP iza starog
'            retExeCode.Add retExeCode.Item(prevIf) & retExeCode.Count + 1, , , prevIf
'            'izbrišemo stari
'            retExeCode.Remove prevIf
'            'brišemo sa stoga prijašnji JP
'            condStog.Remove condStog.Count
            'pokizavac na liniju kako bi se znalo vratiti
            '   kod end if na to i staviti liniju za skok
'            condStog.Add retExeCode.Count
        
        'end if
        ElseIf InStrB(1, strLn, "end if", vbTextCompare) <> 0 Or InStrB(1, strLn, "end if", vbTextCompare) <> 0 Then  'StrComp(Mid$(strLn, 1, Len("end if")), "end if", vbTextCompare) = 0 Or StrComp(Mid$(strLn, 1, Len("endif")), "endif", vbTextCompare) = 0 Then
'            prevIf = condStog.Item(condStog.Count)
'            retExeCode.Add retExeCode.Item(prevIf) & retExeCode.Count + 1, , , prevIf
'            retExeCode.Remove condStog.Item(condStog.Count)
'            condStog.Remove condStog.Count
        
        'if
        ElseIf InStrB(1, strLn, "if", vbTextCompare) <> 0 Then 'StrComp(Mid$(strLn, 1, Len("if")), "if", vbTextCompare) = 0 Then
            
            getIF strLn, mIf
            retExeCode.Add "$tmp_IfSYSVAR_i=" & Replace$(mIf.ifCondition, " ", vbNullString)
            retExeCode.Add "JP:$tmp_IfSYSVAR_i==0:" & getNextIfElseIf(codeHolder, i)
            'pokizavac na liniju kako bi se znalo vratiti
            '   kod end if na to i staviti liniju za skok
            'condStog.Add retExeCode.Count
        
        'else
        ElseIf InStrB(1, strLn, "else", vbTextCompare) <> 0 Then ''StrComp(Mid$(strLn, 1, Len("else")), "else", vbTextCompare) = 0 Then
           
           'prvo stavimo kod koji preskace sve do end if
            '   ukoliko je blok prije izvršen
            retExeCode.Add "JPF:" & getNextIfElseIf(codeHolder, i, True)
        
        'while
        ElseIf InStrB(1, strLn, "while", vbTextCompare) <> 0 Then 'StrComp(Mid$(strLn, 1, Len("while")), "while", vbTextCompare) = 0 Then
            
            getWhile strLn, mWhile
            retExeCode.Add "$varSysWhile" & whileCnt & "=" & Replace$(mWhile.whileCondition, " ", vbNullString)
            retExeCode.Add "JP:$varSysWhile" & whileCnt & "==0:" & getWhileEnd(codeHolder, i)
            
            whileCond.Add mWhile.whileCondition
            whileCnt = whileCnt + 1
            'stavimo pokazivac gdje pocinje while kod
            whileStog.Add retExeCode.Count + 1
           
        'break
        ElseIf InStrB(1, strLn, "break", vbTextCompare) <> 0 Then 'StrComp(Mid$(strLn, 1, Len("while")), "while", vbTextCompare) = 0 Then
            
            'getWhile strLn, mWhile
            'retExeCode.Add "$varSysWhile" & whileCnt & "=" & mWhile.whileCondition
            'MsgBox "BREAK"
            retExeCode.Add "JPF:" & getWhileEnd(codeHolder, i)
            
            'whileCond.Add mWhile.whileCondition
            'whileCnt = whileCnt + 1
            'stavimo pokazivac gdje pocinje while kod
            'whileStog.Add retExeCode.Count + 1
          
        'loop
        ElseIf InStrB(1, strLn, "loop", vbTextCompare) <> 0 Then 'StrComp(Mid$(strLn, 1, Len("loop")), "loop", vbTextCompare) = 0 Then
            
            whileCnt = whileCnt - 1
            retExeCode.Add "$varSysWhile" & whileCnt & "=" & Replace$(whileCond.Item(whileCond.Count), " ", vbNullString)
            retExeCode.Add "JP:$varSysWhile" & whileCnt & "==1:" & whileStog.Item(whileStog.Count)
            
            whileCond.Remove whileCond.Count
            whileStog.Remove whileStog.Count
            
        'matematièke operacije
        ElseIf InStrB(1, strLn, "=") > 0 Then
            
            strLn = Trim$(strLn)
            If right$(strLn, 1) = ";" Then
                strLn = Mid$(strLn, 1, Len(strLn) - 1)
            End If
            retExeCode.Add Replace$(strLn, " ", vbNullString)
            
        'poruka
        ElseIf InStrB(1, strLn, "message", vbTextCompare) > 0 Then
            
            retExeCode.Add strLn
            
        'greška
        ElseIf InStrB(1, strLn, "setError", vbTextCompare) > 0 Then
            
            retExeCode.Add strLn
            
        End If
            
    Next i
    
    Erase tmpVarNms
    Erase tmpVarVals
    Set whileStog = Nothing
    Set whileCond = Nothing
End Sub

Private Function getNextIfElseIf(ByRef codeHolder As Collection, ByVal mLn As Integer, Optional getEndIf As Boolean = False) As Integer
    
    Dim tmpLvl As Integer, strLn As String
    Dim cnt1 As Integer
    Dim i As Integer
    Dim tmpVarNms() As String, tmpVarVals() As String
    
    getNextIfElseIf = 0
    tmpLvl = 0
    
    For i = 1 To codeHolder.Count
    
        strLn = codeHolder.Item(i)
        '
        If InStrB(1, strLn, "return", vbTextCompare) <> 0 Then
            
            getNextIfElseIf = getNextIfElseIf + 1
            
        'message
        ElseIf InStrB(1, strLn, "message", vbTextCompare) <> 0 Then
            
            getNextIfElseIf = getNextIfElseIf + 1
            
        'set error
        ElseIf InStrB(1, strLn, "setError", vbTextCompare) <> 0 Then
            
            getNextIfElseIf = getNextIfElseIf + 1
            
        ElseIf InStrB(1, strLn, "setError", vbTextCompare) <> 0 Then
            
            getNextIfElseIf = getNextIfElseIf + 1
            
        'define
        ElseIf InStrB(1, strLn, "define", vbTextCompare) <> 0 Then
            
            'svaka varijabla u novi red
            getDefineVars strLn, tmpVarNms, tmpVarVals
            getNextIfElseIf = getNextIfElseIf + UBound(tmpVarNms) + 1
            
        'else if
        ElseIf InStrB(1, strLn, "elseif", vbTextCompare) <> 0 Then
            
            getNextIfElseIf = getNextIfElseIf + 3
            
            If i > mLn Then
                
                'ukoliko treba vratiti end if onda
                '   se ovo preskace
                If getEndIf = False Then
                    If tmpLvl = 0 Then
                        getNextIfElseIf = getNextIfElseIf - 1
                        GoTo endF
                    End If
                End If
                
            ElseIf i = mLn Then
                
                tmpLvl = 0
                
            End If
            
        'end if
        ElseIf InStrB(1, strLn, "end if", vbTextCompare) <> 0 Or InStrB(1, strLn, "end if", vbTextCompare) <> 0 Then  'StrComp(Mid$(strLn, 1, Len("end if")), "end if", vbTextCompare) = 0 Or StrComp(Mid$(strLn, 1, Len("endif")), "endif", vbTextCompare) = 0 Then
            
            If i > mLn Then
                
                'getNextIfElseIf = getNextIfElseIf + 1
                If tmpLvl = 0 Then
                    If getEndIf <> True Then
                        getNextIfElseIf = getNextIfElseIf + 1
                    End If
                    GoTo endF
                End If
                tmpLvl = tmpLvl - 1
                
            ElseIf i = mLn Then
                
                tmpLvl = 0
                
            End If
            
        'if
        ElseIf InStrB(1, strLn, "if", vbTextCompare) <> 0 Then 'StrComp(Mid$(strLn, 1, Len("if")), "if", vbTextCompare) = 0 Then
            
            getNextIfElseIf = getNextIfElseIf + 2
            If i > mLn Then
                tmpLvl = tmpLvl + 1
            ElseIf i = mLn Then
                tmpLvl = 0
            End If
            
        'else
        ElseIf InStrB(1, strLn, "else", vbTextCompare) <> 0 Then ''StrComp(Mid$(strLn, 1, Len("else")), "else", vbTextCompare) = 0 Then
           
           getNextIfElseIf = getNextIfElseIf + 2
           If i > mLn Then
                'ukoliko treba vratiti end if onda
                '   se ovo preskace
                If getEndIf = False Then
                    If tmpLvl = 0 Then GoTo endF
                End If
            ElseIf i = mLn Then
                tmpLvl = 0
            End If
            
        'while
        ElseIf InStrB(1, strLn, "while", vbTextCompare) <> 0 Then 'StrComp(Mid$(strLn, 1, Len("while")), "while", vbTextCompare) = 0 Then
            
            getNextIfElseIf = getNextIfElseIf + 2
            
        'break
        ElseIf InStrB(1, strLn, "break", vbTextCompare) <> 0 Then ''StrComp(Mid$(strLn, 1, Len("break")), "break", vbTextCompare) = 0 Then
            
            getNextIfElseIf = getNextIfElseIf + 1
            
        'loop
        ElseIf InStrB(1, strLn, "loop", vbTextCompare) <> 0 Then 'StrComp(Mid$(strLn, 1, Len("loop")), "loop", vbTextCompare) = 0 Then
            
            getNextIfElseIf = getNextIfElseIf + 2
            
        'break
        ElseIf InStrB(1, strLn, "break", vbTextCompare) <> 0 Then ''StrComp(Mid$(strLn, 1, Len("break")), "break", vbTextCompare) = 0 Then
            
            getNextIfElseIf = getNextIfElseIf + 1
            
        'matematièke operacije
        ElseIf InStrB(1, strLn, "=") > 0 Then
            
            getNextIfElseIf = getNextIfElseIf + 1
            
        'function
        ElseIf InStrB(1, strLn, "function", vbTextCompare) <> 0 Then ''StrComp(Mid$(strLn, 1, Len("break")), "break", vbTextCompare) = 0 Then
            
            getNextIfElseIf = getNextIfElseIf + 1
        
        End If
    Next i
endF:
    Erase tmpVarNms
    Erase tmpVarVals
End Function

Private Function getWhileEnd(ByRef codeHolder As Collection, ByVal mLn As Integer) As Integer
    Dim tmpLvl As Integer, strLn As String
    Dim cnt1 As Integer
    Dim i As Integer
    Dim tmpVarNms() As String, tmpVarVals() As String
    
    getWhileEnd = 0
    tmpLvl = 0
    
    For i = 1 To codeHolder.Count
        
        strLn = codeHolder.Item(i)
        '
        If InStrB(1, strLn, "return", vbTextCompare) <> 0 Then
            
            getWhileEnd = getWhileEnd + 1
            
        ElseIf InStrB(1, strLn, "message", vbTextCompare) <> 0 Then
            
            getWhileEnd = getWhileEnd + 1
            
        ElseIf InStrB(1, strLn, "setError", vbTextCompare) <> 0 Then
            
            getWhileEnd = getWhileEnd + 1
            
        'define
        ElseIf InStrB(1, strLn, "define", vbTextCompare) <> 0 Then
            
            'svaka varijabla u novi red
            getDefineVars strLn, tmpVarNms, tmpVarVals
            getWhileEnd = getWhileEnd + UBound(tmpVarNms) + 1
            
        'else if
        ElseIf InStrB(1, strLn, "elseif", vbTextCompare) <> 0 Then
            
            getWhileEnd = getWhileEnd + 3
            
        'end if
        ElseIf InStrB(1, strLn, "end if", vbTextCompare) <> 0 Or InStrB(1, strLn, "end if", vbTextCompare) <> 0 Then  'StrComp(Mid$(strLn, 1, Len("end if")), "end if", vbTextCompare) = 0 Or StrComp(Mid$(strLn, 1, Len("endif")), "endif", vbTextCompare) = 0 Then
            
            getWhileEnd = getWhileEnd + 1
            
        'if
        ElseIf InStrB(1, strLn, "if", vbTextCompare) <> 0 Then 'StrComp(Mid$(strLn, 1, Len("if")), "if", vbTextCompare) = 0 Then
            
            getWhileEnd = getWhileEnd + 2
            
        'else
        ElseIf InStrB(1, strLn, "else", vbTextCompare) <> 0 Then ''StrComp(Mid$(strLn, 1, Len("else")), "else", vbTextCompare) = 0 Then
           
           getWhileEnd = getWhileEnd + 2
           
        'while
        ElseIf InStrB(1, strLn, "while", vbTextCompare) <> 0 Then 'StrComp(Mid$(strLn, 1, Len("while")), "while", vbTextCompare) = 0 Then
            
            getWhileEnd = getWhileEnd + 2
            If i > mLn Then
                tmpLvl = tmpLvl + 1
            ElseIf i = mLn Then
                tmpLvl = 0
            End If
        
        'break
        ElseIf InStrB(1, strLn, "break", vbTextCompare) <> 0 Then ''StrComp(Mid$(strLn, 1, Len("break")), "break", vbTextCompare) = 0 Then
            
            getWhileEnd = getWhileEnd + 1
            
        'loop
        ElseIf InStrB(1, strLn, "loop", vbTextCompare) <> 0 Then 'StrComp(Mid$(strLn, 1, Len("loop")), "loop", vbTextCompare) = 0 Then
            
            If tmpLvl <= 0 Then
                'MsgBox tmpLvl
                getWhileEnd = getWhileEnd + 2
                GoTo endF
            End If
            getWhileEnd = getWhileEnd + 2
            tmpLvl = tmpLvl - 1
            
        'break
        ElseIf InStrB(1, strLn, "break", vbTextCompare) <> 0 Then ''StrComp(Mid$(strLn, 1, Len("break")), "break", vbTextCompare) = 0 Then
            
            getWhileEnd = getWhileEnd + 1
            
        'matematièke operacije
        ElseIf InStrB(1, strLn, "=") > 0 Then
            
            
            getWhileEnd = getWhileEnd + 1
        'function
        ElseIf InStrB(1, strLn, "function", vbTextCompare) <> 0 Then ''StrComp(Mid$(strLn, 1, Len("break")), "break", vbTextCompare) = 0 Then
            
            getWhileEnd = getWhileEnd + 1
        
        End If
    Next i
    
endF:
    Erase tmpVarNms
    Erase tmpVarVals
End Function
