VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTablica 
   AutoRedraw      =   -1  'True
   Caption         =   "#"
   ClientHeight    =   5565
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9495
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   633
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3960
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu buttSaveHtml 
         Caption         =   "Save as HTML"
      End
      Begin VB.Menu buttSaveTxt 
         Caption         =   "Save as TXT"
      End
      Begin VB.Menu c1_mnuFile 
         Caption         =   "-"
      End
      Begin VB.Menu buttExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuPrikaz 
      Caption         =   "View"
      Begin VB.Menu buttPrev 
         Caption         =   "Previous page"
         Shortcut        =   ^P
      End
      Begin VB.Menu buttNext 
         Caption         =   "Next page"
         Shortcut        =   ^N
      End
   End
End
Attribute VB_Name = "frmTablica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mY As Long, maxXChars As Integer, maxYChars As Integer
Private colX As New Collection, colY As New Collection
Private firstItm As Long, showCnt As Long
'
Public Sub addData(ByVal xVal As Double, ByVal yVal As Double)

    colX.Add xVal
    colY.Add yVal
    
End Sub

'
Public Sub Redraw()

    Dim yInc As Long, maxXWid As Long, maxYWid As Long
    Dim i As Integer
    '
    yInc = Me.TextHeight("1") + 10
    'naduži x podatak
    maxXWid = 0
    mY = 0
    maxXChars = 1   '=len("x")
    maxYChars = 5   '=len("f($x)")
    '
    For i = 1 To colX.Count
    
        If maxXWid < Me.TextWidth(Str$(colX.Item(i))) Then
            maxXWid = Me.TextWidth(Str$(colX.Item(i)))
        End If
        
        
        If maxXChars < Len(Str$(colX.Item(i))) Then
            maxXChars = Len(Str$(colX.Item(i)))
        End If
        
        
        If maxYWid < Me.TextWidth(Str$(colY.Item(i))) Then
            maxYWid = Me.TextWidth(Str$(colY.Item(i)))
        End If
        
        
        If maxYChars < Len(Str$(colY.Item(i))) Then
            maxYChars = Len(Str$(colY.Item(i)))
        End If
        
    Next i
    '
    'poèetak tablice
    Me.Line (0, mY)-(maxXWid + maxYWid + 100, mY + yInc), &H80C0FF, BF
    Me.Line (0, mY)-(maxXWid + maxYWid + 100, mY + yInc), vbBlack, B
    
    '
    Me.CurrentY = mY + 5
    Me.CurrentX = 10
    Me.Print "x"
    
    Me.CurrentY = mY + 5
    Me.CurrentX = maxXWid + 50
    Me.Print "f($x)"
    '
    mY = mY + yInc
    showCnt = 0
    
    For i = firstItm To colX.Count
    
        showCnt = showCnt + 1
        'crtamo tablicu
        If i Mod 2 = 0 Then
        
            Me.Line (0, mY)-(maxXWid + maxYWid + 100, mY + yInc), &HC0C0C0, BF
            
        End If
        
        Me.Line (0, mY)-(maxXWid + maxYWid + 100, mY + yInc), vbBlack, B
        '
        '
        Me.CurrentY = mY + 5
        Me.CurrentX = 10
        Me.Print colX.Item(i)
        
        Me.CurrentY = mY + 5
        Me.CurrentX = maxXWid + 50
        Me.Print colY.Item(i)
        '
        mY = mY + yInc
        If mY > Me.ScaleHeight - 20 Then Exit For
        
    Next i
    '
    Me.Line (maxXWid + 25, 0)-(maxXWid + 25, mY), vbBlack
    
End Sub

Public Sub ClearAll()

    Me.Cls
    mY = 0
    firstItm = 1
    Set colX = Nothing
    Set colY = Nothing
    
End Sub


Private Sub buttExit_Click()

    Unload Me
    
End Sub

Private Sub buttNext_Click()

    If showCnt + firstItm < colX.Count Then
    
        firstItm = firstItm + showCnt
        
    End If
    
    Me.Cls
    mY = 0
    Me.Redraw
    
End Sub

Private Sub buttPrev_Click()

    If firstItm > 1 Then
    
        firstItm = firstItm - showCnt
        If firstItm < 1 Then firstItm = 1
        
    End If
    Me.Cls
    mY = 0
    Me.Redraw
    
End Sub

Private Sub buttSaveHtml_Click()
    
    Dim ff As Integer, i As Integer
    Dim cHTML As New clsHTML
    '
    On Error GoTo errH
    CD1.Filter = "HTML datoteke (*.html)|*.html"
    CD1.ShowSave
    
    cHTML.NewRow
    
    For i = 1 To colX.Count
    
        cHTML.NewRow
        cHTML.NewColumn colX.Item(i)
        cHTML.NewColumn colY.Item(i)
        
    Next i
    
    ff = FreeFile
    
    Open CD1.FileName For Output As #ff
    
        Print #ff, cHTML.codeHTML(frmMain.txtFunction.Text)
        
    Close #ff
errH:
    Set cHTML = Nothing
    
End Sub

Private Sub buttSaveTxt_Click()

    Dim ff As Integer, i As Integer
    On Error GoTo errH
    CD1.Filter = "Tekstualne datoteke (*.txt)|*.txt"
    CD1.ShowSave
    
    ff = FreeFile
    
    Open CD1.FileName For Output As #ff
        
        Print #ff, "f(x) = " & frmMain.txtFunction.Text
        Print #ff, vbCrLf
        Print #ff, "x" & Space(maxXChars - 1) & "   |   " & "f($x)"
        Print #ff, String(maxXChars + Len("   "), "-") & "|" & String(maxYChars + Len("   "), "-")
        
        For i = 1 To colX.Count
            
            Print #ff, colX.Item(i) & Space(maxXChars - Len(colX.Item(i))) & "   |   " & colY.Item(i)
        
        Next i
        
    Close #ff
errH:

End Sub


Private Sub Form_Load()

    firstItm = 1
    showCnt = 1
    
End Sub


Private Sub Form_Resize()

    Me.Cls
    mY = 0
    Me.Redraw
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set colX = Nothing
    Set colY = Nothing
    
End Sub
