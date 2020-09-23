Attribute VB_Name = "modStatistic"
Option Explicit
Public Function permutations(ByVal n As Double, ByVal k As Double) As Double
    Dim numb1 As Double, numb2 As Double
    Dim k1 As Integer, z As Integer
    numb1 = CDbl(k)
    numb2 = CDbl(n)
    
    If numb1 <= numb2 Then
        permutations = 1
        z = numb2
        For k1 = 1 To numb1
            permutations = permutations * z
            z = z - 1
        Next k1
    Else
        permutations = 0
    End If
End Function


Public Function combinations(ByVal n As Double, ByVal k As Double) As Double
    If k <= n Then
        combinations = factoriel(n) / (factoriel(k) * factoriel(n - k))
    Else
        combinations = 0
    End If
End Function

Public Function variations(ByVal n As Double, ByVal k As Double) As Double
    If k <= n Then
        'If Fix(n) = n And Fix(k) = k Then
            variations = (factoriel(n) / (factoriel(k) * factoriel(n - k))) * factoriel(k)
        'Else
          '  variations = (functionGamma(n + 1) / (functionGamma(k + 1) * functionGamma(n - k + 1))) * functionGamma(k + 1)
       ' End If
    Else
        variations = 0
    End If
End Function

Public Function factoriel(ByVal num As Double) As Double
    'provjera parametra
    If num < 0 Then
        lastErr = ERR_NegArg
        lastErrNum = ERR_NegArgN
        Exit Function
'    ElseIf num <> Int(num) Then
'        lastErr = ERR_NonIntArg
'        lastErrNum = ERR_NonIntArgN
'        Exit Function
    End If
    
    '
    If num = Fix(num) Then
        
        Dim i As Integer
        factoriel = 1
        
        For i = 1 To num
        
            factoriel = factoriel * i
            
        Next i
        
    Else
    
        factoriel = functionGamma(num + 1)
        
    End If
End Function


Private Function factorielInt(ByVal num As Integer) As Double
    Dim i As Integer
    factorielInt = 1
    For i = 1 To num
        factorielInt = factorielInt * i
    Next i
End Function
