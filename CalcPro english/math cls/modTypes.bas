Attribute VB_Name = "modTypes"

'angle mode
'Public Enum eMode
'    Radian
'    Degree
'End Enum
'complex data type
Public Type Complex
    Re As Double
    Im As Double
End Type
'if(a = b \\ a < 1 \\ (a = 1 && b = 0))
Public Type cmdIF
    'ifStart As Integer
    'ifEnd As Integer
    ifCondition As String
    'ifTrue As String
End Type
'for(i=1;i<=b;step=1)
Public Type cmdFor
    startPT As Double
    endPT As Double
    step As Double
    currPT As Double
    varName As String
    isCompl As Boolean
End Type
'while (a<=b)
Public Type cmdWhile
    startLn As Integer
    whileCondition As String
    mLevel As Integer
End Type





