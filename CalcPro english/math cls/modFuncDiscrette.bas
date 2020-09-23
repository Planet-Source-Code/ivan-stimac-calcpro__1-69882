Attribute VB_Name = "modFuncDiscrete"
Option Explicit

'Binomial cumulative distribution function (CDF).
Public Function BinomCDF(ByVal mX As Double, ByVal n As Integer, ByVal p As Double) As Double
    Dim j As Integer, x1 As Integer
    BinomCDF = 0
    x1 = Int(mX)
    
    If mX < 0 Or n < 0 Then
        lastErr = ERR_BinomNegPar
        lastErrNum = ERR_BinomNegParN
        Exit Function
        
    ElseIf mX > n Then
        lastErr = ERR_BinomArgErr
        lastErrNum = ERR_BinomArgErr
        Exit Function
        
    ElseIf p < 0 Or p > 1 Then
        lastErr = ERR_BinomLastErr
        lastErrNum = ERR_BinomLastErrN
        Exit Function
        
    'ako nije integer
   ' ElseIf Fix(Abs(n)) - Abs(n) <> 0 Then
        
        
    End If
    '
    BinomCDF = 0
    For j = 1 To x1
        'MsgBox (1 - p)
        BinomCDF = BinomCDF + (combinations(n, j) * (p ^ j) * ((1 - p) ^ (n - j)))
    Next j
End Function

'Binomial Probability mass function
Public Function BinomPMF(ByVal mX As Double, ByVal n As Integer, ByVal p As Double) As Double
    Dim j As Integer, x1 As Integer
    BinomPMF = 0
    x1 = Int(mX)

    If mX < 0 Or n < 0 Then
        lastErr = ERR_BinomNegPar
        lastErrNum = ERR_BinomNegParN
        Exit Function
    ElseIf mX > n Then
        lastErr = ERR_BinomArgErr
        lastErrNum = ERR_BinomArgErr
        Exit Function
    ElseIf p < 0 Or p > 1 Then
        lastErr = ERR_BinomLastErr
        lastErrNum = ERR_BinomLastErrN
        Exit Function
    End If
    '
    BinomPMF = combinations(n, mX) * p ^ mX * (1 - p) ^ (n - mX)
End Function
'
'Hypergeometric distribution
'Hypergeometric Probability mass function (pmf)
' k  - is the number of successes in the sample.
' n1 - is the size of the sample
'' m  - is the number of successes in the population.
'' n  - is the population size.
'Public Function HypergeometricCDF(ByVal k As Double, ByVal n1 As Double, ByVal m As Double, ByVal n As Double) As Double
'    If k < 0 Or m < 0 Then
'        lastErr = ERR_hiperGeomPMF
'        lastErrNum = ERR_hiperGeomPMFN
'        Exit Function
'    ElseIf n < 1 Or n1 < 1 Then
'        lastErr = ERR_hiperGeomPMF
'        lastErrNum = ERR_hiperGeomPMFN
'        Exit Function
'    ElseIf n1 < k Or n < m Then
'        lastErr = ERR_hiperGeomPMF
'        lastErrNum = ERR_hiperGeomPMFN
'        Exit Function
'    ElseIf m < k Or n < n1 Then
'        lastErr = ERR_hiperGeomPMF
'        lastErrNum = ERR_hiperGeomPMFN
'        Exit Function
'    End If
'    HypergeometricCDF = combinations(m, k) * combinations(n - m, n1 - k) / combinations(n, n1)
'End Function

'

'Geometric distr
Public Function GeometricPMF(ByVal p As Double, k As Integer) As Double
    
    If p < 0 Or p > 1 Then
        lastErr = ERR_geometricFirst
        lastErrNum = ERR_geometricFirstN
        Exit Function
    ElseIf k < 1 Then
        lastErr = ERR_geometricLast
        lastErrNum = ERR_geometricLastN
        Exit Function
    End If
    GeometricPMF = (1 - p) ^ (k - 1) * p
    
End Function

'Geometric distr
Public Function GeometricCDF(ByVal p As Double, k As Integer) As Double
    
    If p < 0 Or p > 1 Then
        lastErr = ERR_geometricFirst
        lastErrNum = ERR_geometricFirstN
        Exit Function
    ElseIf k < 1 Then
        lastErr = ERR_geometricLast
        lastErrNum = ERR_geometricLastN
        Exit Function
    End If
    GeometricCDF = 1 - (1 - p) ^ k
    
End Function
'
'PoissonPMF distr
'Public Function PoissonPMF(ByVal k As Double, lambda As Integer) As Double
'    If k < 0 Or lambda < 0 Then
'        lastErr = ERR_PoissonNeg
'        lastErrNum = ERR_PoissonNegN
'        Exit Function
'    End If
'    PoissonPMF = Exp(-lambda) * lambda ^ k / factoriel(k)
'End Function
''PoissonPMF distr
'Public Function PoissonCDF(ByVal k As Double, lambda As Integer) As Double
'    Dim i As Integer
'    If k < 0 Or lambda < 0 Then
'        lastErr = ERR_PoissonNeg
'        lastErrNum = ERR_PoissonNegN
'        Exit Function
'    End If
'    PoissonCDF = 0
'    'MsgBox Exp(-lambda)
'    For i = 0 To k
'        PoissonCDF = PoissonCDF + lambda ^ i * Exp(-lambda) / factoriel(k)
'    Next i
'End Function
