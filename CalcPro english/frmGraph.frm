VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGraph 
   Caption         =   "Graph"
   ClientHeight    =   4365
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4860
      Top             =   2220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox pcGraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   91
      TabIndex        =   0
      Top             =   0
      Width           =   1395
   End
   Begin VB.Menu mnuDatoteka 
      Caption         =   "File"
      Begin VB.Menu buttOpen 
         Caption         =   "Open graph"
         Index           =   0
      End
      Begin VB.Menu buttOpen 
         Caption         =   "Add graph from file"
         Index           =   1
      End
      Begin VB.Menu c1_mnuFile 
         Caption         =   "-"
      End
      Begin VB.Menu buttSave 
         Caption         =   "Save graph"
      End
      Begin VB.Menu buttSavePic 
         Caption         =   "Save graph as image"
      End
      Begin VB.Menu c2_mnuFile 
         Caption         =   "-"
      End
      Begin VB.Menu buttExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuPrikaz 
      Caption         =   "View"
      Begin VB.Menu buttLegend 
         Caption         =   "Legend"
      End
      Begin VB.Menu buttFontSize 
         Caption         =   "Legend font size"
         Begin VB.Menu s1 
            Caption         =   "8"
            Index           =   0
         End
         Begin VB.Menu s1 
            Caption         =   "9"
            Index           =   1
         End
         Begin VB.Menu s1 
            Caption         =   "10"
            Index           =   2
         End
         Begin VB.Menu s1 
            Caption         =   "12"
            Index           =   3
         End
         Begin VB.Menu s1 
            Caption         =   "14"
            Index           =   4
         End
      End
      Begin VB.Menu buttFontSkala 
         Caption         =   "Scale font size"
         Begin VB.Menu s2 
            Caption         =   "8"
            Index           =   0
         End
         Begin VB.Menu s2 
            Caption         =   "9"
            Index           =   1
         End
         Begin VB.Menu s2 
            Caption         =   "10"
            Index           =   2
         End
         Begin VB.Menu s2 
            Caption         =   "12"
            Index           =   3
         End
         Begin VB.Menu s2 
            Caption         =   "14"
            Index           =   4
         End
      End
   End
   Begin VB.Menu mnuOpcije 
      Caption         =   "Draw options"
      Begin VB.Menu buttNeCrtajSkokove 
         Caption         =   "Don't draw big changes"
      End
      Begin VB.Menu buttMnuMaxDist 
         Caption         =   "Max y distance between two points "
         Begin VB.Menu buttOdabir 
            Caption         =   "2*maxX"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu buttOdabir 
            Caption         =   "maxX"
            Index           =   1
         End
         Begin VB.Menu buttOdabir 
            Caption         =   "0.5*maxX"
            Index           =   2
         End
         Begin VB.Menu buttOdabir 
            Caption         =   "0.3*maxX"
            Index           =   3
         End
      End
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cGraph As New clsGraph
Private cCalc As New clsCalculator
Private colFuncs As New Collection, colFuncColors As New Collection
Public maxX As Single, maxY As Single, korak As Single

Private Sub buttExit_Click()

    Unload Me
    
End Sub

'prikazivanje legende
Private Sub buttLegend_Click()

    Me.buttLegend.Checked = Not buttLegend.Checked
    Me.Redraw
    
End Sub
'
Private Sub printLegend()

    Dim i As Integer, maxWid As Long
    If Me.buttLegend.Checked = True Then
        'tražimo najdulju f-ju
        maxWid = 0
        
        For i = 1 To colFuncs.Count
            
            If Me.pcGraph.TextWidth(colFuncs.Item(i)) > maxWid Then
                maxWid = Me.pcGraph.TextWidth(colFuncs.Item(i))
            End If
            
        Next i
        '
        Me.pcGraph.Line (Me.pcGraph.ScaleWidth - maxWid - 10, 5)-(Me.pcGraph.ScaleWidth - 5, Me.pcGraph.TextHeight("1") * colFuncs.Count + 5), , B
        
        'ispis legende
        For i = 0 To colFuncs.Count - 1
            
            Me.pcGraph.CurrentY = 5 + Me.pcGraph.TextHeight("1") * i '+ 5 * i
            Me.pcGraph.CurrentX = Me.pcGraph.ScaleWidth - maxWid - 8
            Me.pcGraph.ForeColor = colFuncColors.Item(i + 1)
            Me.pcGraph.Print colFuncs.Item(i + 1)
            
        Next i
    End If
End Sub



Private Sub buttNeCrtajSkokove_Click()
    
    Me.buttNeCrtajSkokove.Checked = Not Me.buttNeCrtajSkokove.Checked
    '
    If Me.buttNeCrtajSkokove.Checked = True Then
        SaveSetting "CalcPro", "Data", "neCrtajSkokove", 1
    Else
        SaveSetting "CalcPro", "Data", "neCrtajSkokove", 0
    End If
    Me.Redraw
    
End Sub

Private Sub buttOdabir_Click(Index As Integer)
    
    Dim i As Integer
    
    For i = 0 To Me.buttOdabir.Count - 1
        
        Me.buttOdabir(i).Checked = False
        
    Next i
    
    Me.buttOdabir(Index).Checked = True
    Me.Redraw
    
End Sub

Private Sub buttOpen_Click(Index As Integer)
    
    On Error GoTo errH
    Dim ff As Integer
    Dim strLn As String
    CD1.Filter = "CalcPro Graph File (*.cpg)|*.cpg"
    CD1.ShowOpen
    
    'poništavanje prijašnjih grafova
    If Index = 0 Then
        
        Set colFuncs = Nothing
        Set colFuncColors = Nothing
        
    End If
    
    ff = FreeFile
    'ucitavanje boja
    Open CD1.FileName For Input As #ff
        
        Do Until EOF(ff)
            
            Line Input #ff, strLn
            
            If left$(strLn, 3) = "FX:" Then
                colFuncs.Add Mid$(strLn, 4)
            ElseIf left$(strLn, 3) = "FC:" Then
                colFuncColors.Add Mid$(strLn, 4)
            End If
            
        Loop
        
    Close #ff
    
    'crtanje f-ja
    Me.Redraw
errH:

End Sub

Private Sub buttSave_Click()
    
    On Error GoTo errH
    Dim ff As Integer, i As Integer
    
    CD1.Filter = "CalcPro Graph File (*.cpg)|*.cpg"
    CD1.ShowSave
    ff = FreeFile
    
    Open CD1.FileName For Output As #ff
        
        For i = 1 To colFuncs.Count
            
            Print #ff, "FX:" & colFuncs.Item(i)
            Print #ff, "FC:" & colFuncColors.Item(i)
            
        Next i
        
    Close #ff
errH:

End Sub

Private Sub buttSavePic_Click()
    
    On Error GoTo errH

    CD1.Filter = "BMP file (*.bmp)|*.bmp"
    CD1.ShowSave
    
    SavePicture Me.pcGraph.Image, CD1.FileName
    
errH:

End Sub

Private Sub Form_Deactivate()

    frmOpcije.SetFocus
    
End Sub

Private Sub Form_Load()

    Dim index1 As Integer
    Me.pcGraph.left = 0
    Me.pcGraph.top = 0
    Set cGraph.PicBox = Me.pcGraph
    
    cGraph.BackColor = vbWhite
    cGraph.GraphColor = vbBlack
    cGraph.yMax = maxY
    cGraph.xMax = maxX
    cGraph.xScale = cGraph.xMax / 10
    cGraph.yScale = cGraph.yMax / 10
    
    index1 = GetSetting("CalcPro", "Data", "legendFontSizei", 0)
    Me.s1(index1).Checked = True
    
    index1 = GetSetting("CalcPro", "Data", "scaleFontSizei", 0)
    Me.s2(index1).Checked = True
    '
    buttNeCrtajSkokove.Checked = GetSetting("CalcPro", "Data", "neCrtajSkokove", 0)
    
End Sub



Private Sub Form_Resize()

    On Error GoTo errH
    Dim i As Integer
    Me.pcGraph.Width = Me.Width - 100
    Me.pcGraph.Height = Me.Height - 800
    cGraph.drawGraph
    
    Me.pcGraph.FontSize = Val(GetSetting("CalcPro", "Data", "legendFontSize", 8))
    printLegend
    
    Me.pcGraph.FontSize = Val(GetSetting("CalcPro", "Data", "scaleFontSize", 8))
    
    For i = 1 To colFuncs.Count
        
        drawFunction colFuncs.Item(i), colFuncColors.Item(i)
        
    Next i
    
errH:
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set cGraph = Nothing
    Set colFuncs = Nothing
    Set colFuncColors = Nothing
    Set cCalc = Nothing
    
End Sub

'

Private Sub drawFunction(ByVal strFunction As String, ByVal Color1 As OLE_COLOR)

    Dim mY1 As Single, mY2 As Single
    Dim mX1 As Single
    Dim i As Single, maxDist As Single
    
    '
    If Me.buttOdabir(0).Checked = True Then
        maxDist = maxX * 2
        
    ElseIf Me.buttOdabir(1).Checked = True Then
        maxDist = maxX * 2
        
    ElseIf Me.buttOdabir(2).Checked = True Then
        maxDist = maxX / 2
        
    ElseIf Me.buttOdabir(3).Checked = True Then
        maxDist = maxX / 3
        
    End If
    
    i = -maxX
    frmMain.clearError
    mX1 = i
    
    Do While i <= maxX
    
        mY1 = frmMain.getValue(strFunction, mX1)
        i = i + korak
        'ako nema greške
        If frmMain.getLastErr = 0 Then
            
            mY2 = frmMain.getValue(strFunction, i)
            
            If frmMain.getLastErr = 0 Then
                If Abs(mY1 - mY2) > maxDist And Me.buttNeCrtajSkokove.Checked = True Then
                    i = i - korak
                Else
                    cGraph.DrawLine mX1, mY1, i, mY2, Color1
                End If
            End If
            
        End If
        
        mX1 = i
        i = i + korak
        frmMain.clearError
        
    Loop
    
End Sub


Public Sub AddFunction(ByVal strFunction As String, ByVal Color1 As OLE_COLOR)

    colFuncs.Add strFunction
    colFuncColors.Add Color1
    
End Sub

Public Function getFNum() As Integer

    getFNum = colFuncs.Count
    
End Function

Public Sub Redraw()

    cGraph.yMax = maxY
    cGraph.xMax = maxX
    cGraph.xScale = cGraph.xMax / 10
    cGraph.yScale = cGraph.yMax / 10
    Form_Resize
    
End Sub


Private Sub s1_Click(Index As Integer)

    Dim i As Integer
    
    For i = 0 To s1.Count - 1
        s1(i).Checked = False
    Next i
    
    SaveSetting "CalcPro", "Data", "legendFontSize", Me.s1(Index).Caption
    SaveSetting "CalcPro", "Data", "legendFontSizei", Index
    
    s1(Index).Checked = True
    
    Me.Redraw
    
End Sub


Private Sub s2_Click(Index As Integer)

    Dim i As Integer
    
    For i = 0 To s2.Count - 1
        s2(i).Checked = False
    Next i
    
    SaveSetting "CalcPro", "Data", "scaleFontSize", Me.s2(Index).Caption
    SaveSetting "CalcPro", "Data", "scaleFontSizei", Index
    
    s2(Index).Checked = True
    
    Me.Redraw
    
End Sub
