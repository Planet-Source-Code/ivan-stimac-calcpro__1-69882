Attribute VB_Name = "modFuncsNumb"
Option Explicit

Public Function cFrac(ByRef z As Complex) As Complex

    Dim sign As Integer
    Dim str1 As String
    
    sign = 1
    If z.Re < 0 Then
        sign = -1
    End If
    
    z.Re = Abs(z.Re)
    
    str1 = Str$(z.Re)
    
    If InStrB(1, str1, ",") Then
    
        str1 = Replace$(str1, ",", ".")
        
    End If
    
    If InStrB(1, str1, ".") Then
    
        str1 = "0." & Mid$(str1, InStr(1, str1, ".") + 1)
        cFrac.Re = Val(str1) * sign
        
    Else
    
        cFrac.Re = 0
        
    End If
    
    
    sign = 1
    If z.Im < 0 Then
    
        sign = -1
        
    End If
    
    
    str1 = Str$(z.Im)
    
    If InStrB(1, str1, ",") Then
    
        str1 = Replace$(str1, ",", ".")
        
    End If
    If InStrB(1, str1, ".") Then
    
        str1 = "0." & Mid$(str1, InStr(1, str1, ".") + 1)
        cFrac.Im = Val(str1) * sign
        
    Else
    
        cFrac.Im = 0
        
    End If
End Function
