Attribute VB_Name = "modFuncNumbConvert"
Option Explicit

'=====================================================================
'=  Dec2Bin        = ------------------------------------------------=
'=====================================================================
Public Function Dec2Bin(ByVal decVal As Double) As String
    
    Dim mBin(9) As String, i As Integer, sign As Integer
    Dim tmpDec As Long
    
    
    'check for errors
    If Fix(Abs(decVal)) - Abs(decVal) <> 0 Then
        
        lastErr = "Dec2Bin" & ERR_NonIntVal
        lastErrNum = ERR_NonIntValN
        Exit Function
        
    ElseIf decVal < -512 Or decVal > 511 Then
        
        lastErr = "Dec2Bin" & Replace$(ERR_OutsideInterval, "[i]", "[-512, 511]")
        lastErrNum = ERR_OutsideIntervalN
        Exit Function
        
    End If
    
    
    
    'find sign
    If decVal < 0 Then
        sign = -1
        decVal = Abs(decVal)
        
    Else
        sign = 0
        
    End If
    
    '
    tmpDec = CLng(decVal)
    i = 0
    
    'find binary value for unsigned value
    Do While tmpDec >= 1
        
        i = i + 1
        mBin(10 - i) = tmpDec Mod 2
        tmpDec = tmpDec \ 2
        
    Loop
    
    
    'for negative value find 2's complement
    If sign = -1 Then
        
        '1's complement
        For i = 0 To 9
            
            If mBin(9 - i) = "1" Then
                mBin(9 - i) = "0"
            Else
                mBin(9 - i) = "1"
            End If
            
        Next i
        
        
        '2's complement (add 1)
        For i = 0 To 9
        
            If mBin(9 - i) = "1" Then
                mBin(9 - i) = "0"
            Else
                mBin(9 - i) = "1"
                Exit For
            End If
            
        Next i
    End If
    
    
    Dec2Bin = Join$(mBin, vbNullString)
End Function

'=====================================================================
'=  Bin2Dec        = ------------------------------------------------=
'=====================================================================
Public Function Bin2Dec(ByVal binVal As String) As Long
    Dim sign As Integer, i As Integer, len1 As Integer, val1 As Integer
    Dim tmpVal As Double
    
    
    sign = 1
    tmpVal = 0
    len1 = Len(binVal)
    
    
    If len1 > 10 Or len1 < 1 Then
    
        lastErr = ERR_InvalidBinLen
        lastErrNum = ERR_InvalidBinLenN
        Exit Function
        
    End If
    
    
    
    'if len is 10 and first char = 1
    '   then number is negative
    
    If len1 = 10 And Val(Mid$(binVal, 1, 1)) = 1 Then
    
        sign = -1
        
    End If
    
    
    For i = 1 To len1
       
        Select Case Mid$(binVal, i, 1)
            
            Case "0", "1"
                
                val1 = Val(Mid$(binVal, i, 1))
                tmpVal = tmpVal + val1 * 2 ^ (len1 - i)
            
            
            'error
            Case Else
            
                lastErr = ERR_InvalidBinVal
                lastErrNum = ERR_InvalidBinValN
                Exit Function
                
        End Select
        
    Next i
    

    
    'if we have negative value, then
    '   find this value (for negative numbers
    '   FFFFFFFF is -1, FFFFFFFE is -2 ... so
    '   we for unsigned value substract min value
    '   (for long data type) * 2
    If sign = -1 Then
    
        tmpVal = tmpVal - 512 * 2
        
    End If
    
    Bin2Dec = CLng(tmpVal)
End Function

'=====================================================================
'=  Dec2Hex        = ------------------------------------------------=
'=====================================================================
Public Function Dec2Hex(ByVal decVal As Double) As String
    
    'check for errors
    If Fix(Abs(decVal)) - Abs(decVal) <> 0 Then
        
        lastErr = "Dec2Hex" & ERR_NonIntVal
        lastErrNum = ERR_NonIntValN
        Exit Function
        
    ElseIf decVal < -2147483648# Or decVal > 2147483647 Then
        
        lastErr = "Dec2Hex" & Replace$(ERR_OutsideInterval, "[i]", "[-2147483648, 2147483647]")
        lastErrNum = ERR_OutsideIntervalN
        Exit Function
        
    End If
    
    Dec2Hex = Hex(decVal)
    
    
    If decVal < 0 Then
        
        If Len(Dec2Hex) < 8 Then
        
            Dec2Hex = String(Len(Dec2Hex) - 8, "F") & Dec2Hex
        End If
        
    End If
    
End Function

'=====================================================================
'=  Dec2Oct        = ------------------------------------------------=
'=====================================================================
Public Function Dec2Oct(ByVal decVal As Double) As String
    
    'check for errors
    If Fix(Abs(decVal)) - Abs(decVal) <> 0 Then
        
        lastErr = "Dec2Oct" & ERR_NonIntVal
        lastErrNum = ERR_NonIntValN
        Exit Function
        
    ElseIf decVal < -536870912 Or decVal > 536870911 Then
        
        lastErr = "Dec2Hex" & Replace$(ERR_OutsideInterval, "[i]", "[-536870912,536870912]")
        lastErrNum = ERR_OutsideIntervalN
        Exit Function
        
    End If
    
    Dec2Oct = Oct(decVal)
    
    'if we have negative value
    If decVal < 0 Then
        
        'add chars
        If Len(Dec2Oct) < 10 Then
        
            Dec2Oct = String(Len(Dec2Oct) - 8, "7") & Dec2Oct
            
        'remove first char (usualy 3)
        ElseIf Len(Dec2Oct) = 11 Then
            
            Dec2Oct = Mid$(Dec2Oct, 2)
            
        End If
        
    End If
    
End Function

'=====================================================================
'=  hex2dec        = ------------------------------------------------=
'=====================================================================
Public Function Hex2Dec(ByVal hexVal As String) As Long
    Dim sign As Integer, i As Integer, len1 As Integer, val1 As Integer
    Dim tmpVal As Double
    
    
    sign = 1
    tmpVal = 0
    len1 = Len(hexVal)
    
    
    If len1 > 8 Or len1 < 1 Then
    
        lastErr = ERR_InvalidHexLen
        lastErrNum = ERR_InvalidHexLenN
        Exit Function
        
    End If
    
    
    
    'if len is 8 and first char >= 8
    '   then number is negative
    
    If len1 = 8 And (Val(Mid$(hexVal, 1, 1)) > 1 Or IsNumeric(Mid$(hexVal, 1, 1)) = False) Then
    
        sign = -1
        
    End If
    
    
    For i = 1 To len1
       
        Select Case Mid$(hexVal, i, 1)
            
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                
                val1 = Val(Mid$(hexVal, i, 1))
                tmpVal = tmpVal + val1 * 16 ^ (len1 - i)
            
            Case "A"
                
                tmpVal = tmpVal + 10 * 16 ^ (len1 - i)
                
            Case "B"
            
                tmpVal = tmpVal + 11 * 16 ^ (len1 - i)
                
            Case "C"
            
                tmpVal = tmpVal + 12 * 16 ^ (len1 - i)
            
            Case "D"
                
                tmpVal = tmpVal + 13 * 16 ^ (len1 - i)
                
            Case "E"
            
                tmpVal = tmpVal + 14 * 16 ^ (len1 - i)
                
            Case "F"
            
                tmpVal = tmpVal + 15 * 16 ^ (len1 - i)
            
            'error
            Case Else
            
                lastErr = ERR_InvalidHexVal
                lastErrNum = ERR_InvalidHexValN
                Exit Function
                
        End Select
        
    Next i
    

    
    'if we have negative value, then
    '   find this value (for negative numbers
    '   FFFFFFFF is -1, FFFFFFFE is -2 ... so
    '   we for unsigned value substract min value
    '   (for long data type) * 2
    If sign = -1 Then
    
        tmpVal = tmpVal - 2147483648# * 2
        
    End If
    
    Hex2Dec = CLng(tmpVal)
End Function

'=====================================================================
'=  Oct2Dec        = ------------------------------------------------=
'=====================================================================
Public Function Oct2Dec(ByVal octVal As String) As Long
    
    Dim sign As Integer, i As Integer, len1 As Integer, val1 As Integer
    Dim tmpVal As Double
    
    
    sign = 1
    tmpVal = 0
    len1 = Len(octVal)
    
    
    If len1 > 10 Or len1 < 1 Then
    
        lastErr = ERR_InvalidOctLen
        lastErrNum = ERR_InvalidOctLenN
        Exit Function
        
    End If
    
    
    
    'if len is 8 and first char >= 8
    '   then number is negative
    
    If len1 = 10 And Val(octVal) > 7777777777# / 2 Then
    
        sign = -1
        
    End If
    
    
    For i = 1 To len1
       
        Select Case Mid$(octVal, i, 1)
            
            Case "0", "1", "2", "3", "4", "5", "6", "7"
                
                val1 = Val(Mid$(octVal, i, 1))
                tmpVal = tmpVal + val1 * 8 ^ (len1 - i)
            
            'error
            Case Else
            
                lastErr = ERR_InvalidOctVal
                lastErrNum = ERR_InvalidOctVal
                Exit Function
                
        End Select
        
    Next i
    

    
    'if we have negative value, then
    '   find this value
    If sign = -1 Then
    
        tmpVal = tmpVal - 268435456 * 4
        
    End If
    
    Oct2Dec = CLng(tmpVal)
    
End Function


'
Public Function Oct2Bin(ByVal octVal As Double) As String

    Oct2Bin = Dec2Bin(Oct2Dec(octVal))
    
End Function
'
Public Function Oct2Hex(ByVal octVal As Double) As String

    Oct2Hex = Dec2Hex(Oct2Dec(octVal))
    
End Function

Public Function Hex2Bin(ByVal hexVal As String) As String

    Hex2Bin = Dec2Bin(Hex2Dec(hexVal))
    
End Function
'
Public Function Hex2Oct(ByVal hexVal As String) As Double

    Hex2Oct = Dec2Oct(Hex2Dec(hexVal))
    
End Function
'
Public Function Bin2Hex(ByVal binVal As String) As String
    
    Bin2Hex = Dec2Hex(Bin2Dec(binVal))
    
End Function
'
Public Function Bin2Oct(ByVal bin As String) As String
    
    Bin2Oct = Dec2Oct(Bin2Dec(bin))
    
End Function
