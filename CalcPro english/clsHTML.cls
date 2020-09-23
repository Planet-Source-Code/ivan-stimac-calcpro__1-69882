VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private strHTML As String, rowCnt As Integer
'
Public Sub NewRow()

    If rowCnt < 1 Then
        strHTML = strHTML & "<tr>" & vbCrLf
        
    Else
        strHTML = strHTML & "</tr><tr>" & vbCrLf
        
    End If
    rowCnt = rowCnt + 1
    
End Sub
'
Public Sub NewColumn(ByVal colData As String)

    strHTML = strHTML & "   <td class = 'cTable'>" & colData & "</td>" & vbCrLf
    
End Sub
'
Public Property Get codeHTML(ByVal fName As String) As String

    codeHTML = "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>"
codeHTML = "<html xmlns='http://www.w3.org/1999/xhtml'><head><meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1' /><title>$function</title>" & vbCrLf & _
    "<style type='text/css'>" & vbCrLf & _
    ".cTable {border: 1px solid #006699;}" & vbCrLf & _
    ".calcPro {font-size: large; color: #006699; font-weight: bold;}" & vbCrLf & _
    "body,td,th { font-family: Courier New, Courier, monospace;}" & vbCrLf & _
    "body {background-color: #CCCCCC;}" & vbCrLf & _
    "</style></head><body>" & vbCrLf & _
    "<p class='calcPro'>CalcPro</p><p>f($x) = $function</p>" & vbCrLf & _
    "<table width='100%' border='0' cellpadding='2' cellspacing='0'>" & vbCrLf & _
    "<tr><td bgcolor='#006699' class='cTable'>x</td><td bgcolor='#006699' class='cTable'>f(x)</td></tr>" & vbCrLf & _
    "$rows</table></body></html>"
    
    codeHTML = Replace$(codeHTML, "$function", fName)
    codeHTML = Replace$(codeHTML, "$rows", strHTML)
    
End Property

Private Sub Class_Initialize()

    strHTML = vbNullString
    rowCnt = 0
    
End Sub
