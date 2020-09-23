Attribute VB_Name = "modWriteHTML"
Option Explicit
Private strHTML As String, rowCnt As Integer

Public Sub NewRow()
    If rowCnt < 1 Then
        strHTML = strHTML & "<tr>" & vbCrLf
    Else
        strHTML = strHTML & "</tr><tr>" & vbCrLf
    End If
End Sub

Public Sub NewColumn(ByVal colData As String)
    strHTML = strHTML & "<td>" & colData & "</td>"
End Sub

Public Property Get codeHTML() As String
    
End Property

