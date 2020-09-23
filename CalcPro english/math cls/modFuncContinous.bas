Attribute VB_Name = "modFuncContinous"
Option Explicit

'Chi-square
'Chi-square PDF
Public Function ChiSquarePDF(ByVal mX As Double, ByVal k As Double) As Double
    If mX < 0 Then
        lastErr = ERR_ChiSquareNeg
        lastErrNum = ERR_ChiSquareNegN
        Exit Function
    ElseIf k <= 0 Then
        lastErr = ERR_ChiSquareSecArg
        lastErrNum = ERR_ChiSquareSecArgN
        Exit Function
    End If
    ChiSquarePDF = 0.5 ^ (k / 2) * mX ^ (k / 2 - 1) * Exp(-mX / 2) / functionGamma(k / 2)
End Function
'Chi-square CDF
Public Function ChiSquareCDF(ByVal mX As Double, ByVal k As Double) As Double
    If mX < 0 Then
        lastErr = ERR_ChiSquareNeg
        lastErrNum = ERR_ChiSquareNegN
        Exit Function
    ElseIf k <= 0 Then
        lastErr = ERR_ChiSquareSecArg
        lastErrNum = ERR_ChiSquareSecArgN
        Exit Function
    End If
    ChiSquareCDF = functionGammaIncLwr(k / 2, mX / 2) / functionGamma(k / 2)
End Function
'
'Exponential distribution
'Exponential PDF
Public Function ExponentialPDF(ByVal mX As Double, ByVal lambda As Double) As Double
    If mX < 0 Or lambda < 0 Then
        lastErr = ERR_ExpNeg
        lastErrNum = ERR_ExpNegN
        Exit Function
    End If
    If mX >= 0 Then
        ExponentialPDF = lambda * Exp(-lambda * mX)
    Else
        ExponentialPDF = 0
    End If
End Function
'Exponential CDF
Public Function ExponentialCDF(ByVal mX As Double, ByVal lambda As Double) As Double
    If mX < 0 Or lambda < 0 Then
        lastErr = ERR_ExpNeg
        lastErrNum = ERR_ExpNegN
        Exit Function
    End If
    If mX >= 0 Then
        ExponentialCDF = 1 - Exp(-lambda * mX)
    Else
        ExponentialCDF = 0
    End If
End Function
'
'F-distribution
'
'F-distribution PDF
Public Function FisherPDF(ByVal mX As Double, ByVal d1 As Double, ByVal d2 As Double) As Double
    Dim num1 As Double
    If d1 <= 0 Or d2 <= 0 Or mX <= 0 Then
        lastErr = ERR_FisherSecLastArg
        lastErrNum = ERR_FisherSecLastArgN
        Exit Function
'    ElseIf mX < 0 Then
'        lastErr = ERR_FisherFirstNeg
'        lastErrNum = ERR_FisherFirstNegN
'        Exit Function
    End If
    num1 = (d1 * mX) ^ d1 * d2 ^ d2 / ((d1 * mX + d2) ^ (d1 + d2))
    FisherPDF = Sqr(num1) / (mX * functionBeta(d1 / 2, d2 / 2))
End Function
'F-distribution PDF
Public Function FisherCDF(ByVal mX As Double, ByVal d1 As Double, ByVal d2 As Double) As Double
    Dim num1 As Double
    If d1 <= 0 Or d2 <= 0 Or mX <= 0 Then
        lastErr = ERR_FisherSecLastArg
        lastErrNum = ERR_FisherSecLastArgN
        Exit Function
'    ElseIf mX < 0 Then
'        lastErr = ERR_FisherFirstNeg
'        lastErrNum = ERR_FisherFirstNegN
'        Exit Function
    End If
   'MsgBoxs d1 * mX / (d1 * mX + d2) & vbCrLf & d1 / 2 & vbCrLf & d2 / 2
    FisherCDF = functionBetaInc(d1 * mX / (d1 * mX + d2), d1 / 2, d2 / 2)
End Function
'
'normal distr
Public Function NormalPDF(ByVal mX As Double, ByVal m As Double, ByVal s As Double) As Double
    If s <= 0 Then
        lastErr = ERR_NormlaLastArg
        lastErrNum = ERR_NormlaLastArgN
        Exit Function
    End If
    NormalPDF = Exp(-((mX - m) ^ 2 / (2 * s ^ 2))) / (s * Sqr(2 * PI))
End Function
'normal distr
Public Function NormalCDF(ByVal mX As Double, ByVal m As Double, ByVal s As Double) As Double
    If s <= 0 Then
        lastErr = ERR_NormlaLastArg
        lastErrNum = ERR_NormlaLastArgN
        Exit Function
    End If
    NormalCDF = 0.5 * (1 + functionErf((mX - m) / (s * Sqr(2))))
End Function
