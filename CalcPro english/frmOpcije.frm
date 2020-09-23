VERSION 5.00
Begin VB.Form frmOpcije 
   BorderStyle     =   0  'None
   Caption         =   "frmOpcije"
   ClientHeight    =   9405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8505
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
   ScaleHeight     =   9405
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   1200
      Top             =   1860
   End
   Begin VB.Frame frmAbout 
      BorderStyle     =   0  'None
      Caption         =   "Variable"
      Height          =   1515
      Left            =   0
      TabIndex        =   34
      Top             =   6900
      Width           =   8475
      Begin VB.Label Label22 
         Caption         =   "Web: open webpage"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   3840
         TabIndex        =   40
         Top             =   360
         Width           =   4575
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   3
         Left            =   3540
         Picture         =   "frmOpcije.frx":0000
         Top             =   720
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   2
         Left            =   3540
         Picture         =   "frmOpcije.frx":0062
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label23 
         Caption         =   "Contact: ivan.stimac@po.htnet.hr"
         Height          =   255
         Left            =   3840
         TabIndex        =   39
         Top             =   720
         Width           =   3675
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   1
         Left            =   300
         Picture         =   "frmOpcije.frx":00C4
         Top             =   720
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   0
         Left            =   300
         Picture         =   "frmOpcije.frx":0126
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label21 
         Caption         =   "Version: 1.0"
         Height          =   255
         Left            =   600
         TabIndex        =   38
         Top             =   720
         Width           =   2475
      End
      Begin VB.Label Label17 
         Caption         =   "Name: CalcPro"
         Height          =   255
         Left            =   600
         TabIndex        =   37
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label18 
         Caption         =   "About application:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   60
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   3
         X1              =   3180
         X2              =   3180
         Y1              =   60
         Y2              =   1170
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         X1              =   3150
         X2              =   3150
         Y1              =   60
         Y2              =   1190
      End
      Begin VB.Label Label14 
         Caption         =   "Contact:"
         Height          =   255
         Left            =   3360
         TabIndex        =   35
         Top             =   60
         Width           =   1695
      End
   End
   Begin VB.Frame frmRacun 
      BorderStyle     =   0  'None
      Caption         =   "Variable"
      Height          =   1395
      Left            =   0
      TabIndex        =   24
      Top             =   4560
      Width           =   8475
      Begin VB.CommandButton buttCreateTable 
         Caption         =   "Print table"
         Height          =   675
         Left            =   4680
         Picture         =   "frmOpcije.frx":0188
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   480
         Width           =   1995
      End
      Begin VB.TextBox txtXMin 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1620
         TabIndex        =   27
         Text            =   "-10"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtXMax 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2700
         TabIndex        =   26
         Text            =   "10"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtKorakT 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1620
         TabIndex        =   25
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Print:"
         Height          =   255
         Left            =   4680
         TabIndex        =   33
         Top             =   60
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "to"
         Height          =   435
         Left            =   2340
         TabIndex        =   31
         Top             =   420
         Width           =   675
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   4530
         X2              =   4530
         Y1              =   60
         Y2              =   1190
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   2
         X1              =   4565
         X2              =   4565
         Y1              =   60
         Y2              =   1170
      End
      Begin VB.Label Label20 
         Caption         =   "Table options:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   60
         Width           =   2175
      End
      Begin VB.Label Label19 
         Caption         =   "For $x from"
         Height          =   435
         Left            =   300
         TabIndex        =   29
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Step:"
         Height          =   255
         Left            =   300
         TabIndex        =   28
         Top             =   840
         Width           =   1755
      End
   End
   Begin VB.Frame frmCrtanje 
      BorderStyle     =   0  'None
      Caption         =   "Variable"
      Height          =   1515
      Left            =   0
      TabIndex        =   9
      Top             =   2580
      Width           =   8475
      Begin VB.PictureBox pcBoja 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3780
         ScaleHeight     =   285
         ScaleWidth      =   345
         TabIndex        =   23
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Draw function"
         Height          =   675
         Left            =   4680
         Picture         =   "frmOpcije.frx":11CA
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   480
         Width           =   1995
      End
      Begin VB.TextBox txtKorak 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2700
         TabIndex        =   18
         Text            =   "0.1"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtMaxY 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1620
         TabIndex        =   16
         Text            =   "10"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtMaxX 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   540
         TabIndex        =   12
         Text            =   "10"
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Color:"
         Height          =   255
         Left            =   3660
         TabIndex        =   22
         Top             =   420
         Width           =   795
      End
      Begin VB.Label Label7 
         Caption         =   "Draw:"
         Height          =   255
         Left            =   4680
         TabIndex        =   21
         Top             =   60
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Step:"
         Height          =   255
         Left            =   2460
         TabIndex        =   19
         Top             =   420
         Width           =   1755
      End
      Begin VB.Label Label10 
         Caption         =   "±"
         Height          =   195
         Left            =   1440
         TabIndex        =   17
         Top             =   780
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "±"
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   780
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Y scale:"
         Height          =   255
         Left            =   1380
         TabIndex        =   13
         Top             =   420
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "X scale:"
         Height          =   435
         Left            =   300
         TabIndex        =   11
         Top             =   420
         Width           =   2355
      End
      Begin VB.Label Label4 
         Caption         =   "Graph options:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   60
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   4565
         X2              =   4565
         Y1              =   60
         Y2              =   1170
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   4530
         X2              =   4530
         Y1              =   60
         Y2              =   1190
      End
   End
   Begin VB.Frame frmVarijable 
      BorderStyle     =   0  'None
      Caption         =   "Variable"
      Height          =   1515
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8355
      Begin VB.CommandButton buttSaveVar 
         Caption         =   "Save"
         Height          =   675
         Left            =   6000
         Picture         =   "frmOpcije.frx":220C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Width           =   1275
      End
      Begin VB.TextBox txtVarVal 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5160
         TabIndex        =   7
         Text            =   "0"
         Top             =   540
         Width           =   675
      End
      Begin VB.TextBox txtVarName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4200
         TabIndex        =   5
         Text            =   "$x"
         Top             =   540
         Width           =   675
      End
      Begin VB.ListBox lstVars 
         Appearance      =   0  'Flat
         Height          =   870
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2355
      End
      Begin VB.TextBox txtValue 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2700
         TabIndex        =   1
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "="
         Height          =   255
         Left            =   4980
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "New variable"
         Height          =   255
         Left            =   4080
         TabIndex        =   4
         Top             =   60
         Width           =   2295
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   3930
         X2              =   3930
         Y1              =   60
         Y2              =   1190
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   3960
         X2              =   3960
         Y1              =   60
         Y2              =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Value:"
         Height          =   255
         Left            =   2580
         TabIndex        =   3
         Top             =   60
         Width           =   1515
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   1860
   End
   Begin VB.Label Label9 
      Caption         =   "±"
      Height          =   195
      Left            =   1920
      TabIndex        =   15
      Top             =   3120
      Width           =   255
   End
End
Attribute VB_Name = "frmOpcije"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_NORMAL = 1
Private Const mSpeed As Integer = 1000

Private mX As Long, mY As Long, num1 As Long, num2 As Long

Private Sub buttCreateTable_Click()
    
    If IsNumeric(Me.txtXMin.Text) = False Or InStrB(1, Me.txtXMin.Text, ",") <> 0 Then
    
        MsgBox "Error: start X must be number!", vbInformation
        Exit Sub
    
    ElseIf IsNumeric(Me.txtXMax.Text) = False Or InStrB(1, Me.txtXMax.Text, ",") <> 0 Then
    
        MsgBox "Error: end X must be numberj!", vbInformation
        Exit Sub
    ElseIf IsNumeric(Me.txtKorakT.Text) = False Or InStrB(1, Me.txtKorakT.Text, ",") <> 0 Then
    
        MsgBox "Error: 'Step' must be number!", vbInformation
        Exit Sub
        
    End If
    
    
    Dim i As Single, val1 As Double
    Load frmTablica
    frmTablica.Caption = frmMain.txtFunction.Text
    frmTablica.ClearAll
    
    '
    For i = Val(Me.txtXMin.Text) To Val(Me.txtXMax.Text) Step Val(Me.txtKorakT.Text)
        
        val1 = frmMain.getValue(frmMain.txtFunction.Text, i)
        frmTablica.addData i, val1
        
    Next i
    '
    frmTablica.Redraw
    DoEvents
    frmTablica.Show
    
End Sub

'
Private Sub buttSaveVar_Click()

    frmMain.saveVar Me.txtVarName.Text, Me.txtVarVal.Text
    frmMain.loadVars
    
End Sub

'crtanje funkcije
Private Sub Command2_Click()


    If IsNumeric(Me.txtKorak.Text) = False Or InStrB(1, Me.txtKorak.Text, ",") <> 0 Then
    
        MsgBox "Error: 'Step' must be number!", vbInformation
        Exit Sub
    
    ElseIf IsNumeric(Me.txtMaxY.Text) = False Or InStrB(1, Me.txtMaxY.Text, ",") <> 0 Then
    
        MsgBox "Error: 'Y scale' must be number!", vbInformation
        Exit Sub
        
    ElseIf IsNumeric(Me.txtMaxX.Text) = False Or InStrB(1, Me.txtMaxX.Text, ",") <> 0 Then
    
        MsgBox "Error: 'X scale' must be number!", vbInformation
        Exit Sub
        
    End If
    
    If frmGraph.getFNum < 1 Then
    
        Load frmGraph
        
    End If
    
    'postavljanje paremetara
    frmGraph.maxX = Val(Me.txtMaxX.Text)
    frmGraph.maxY = Val(Me.txtMaxY.Text)
    frmGraph.korak = Val(Me.txtKorak.Text)
    frmGraph.AddFunction frmMain.txtFunction.Text, Me.pcBoja.BackColor
    frmGraph.Redraw
    
    frmGraph.Show
    
End Sub

'
Private Sub Form_Load()

    Me.Height = 1215
    
    Me.frmCrtanje.top = 0
    Me.frmVarijable.top = 0
    Me.frmRacun.top = 0
    Me.frmAbout.top = 0
    
    num1 = 0
    
End Sub

Private Sub Label22_Click()
    ShellExecute hwnd, "open", "http://calcpro.scienceontheweb.net/english", vbNullString, vbNullString, SW_NORMAL
End Sub

Private Sub lstVars_Click()

    Me.txtValue.Text = frmMain.getVarValue(Me.lstVars.List(Me.lstVars.ListIndex))
    
End Sub

'odabir boje
Private Sub pcBoja_Click()

    frmColor.Show vbModal
    
End Sub

Private Sub Timer1_Timer()
        
    If Me.left + mSpeed < frmMain.left Then
        Me.left = Me.left + mSpeed
        
    ElseIf Me.left - mSpeed > frmMain.left Then
        Me.left = Me.left - mSpeed
        
    Else
        Me.left = frmMain.left
        
    End If
    
    
    
    If Me.top + mSpeed < frmMain.top + frmMain.Height Then
        Me.top = Me.top + mSpeed
        
    ElseIf Me.top - mSpeed > frmMain.top + frmMain.Height Then
        Me.top = Me.top - mSpeed
        
    Else
        Me.top = frmMain.top + frmMain.Height
        
    End If
End Sub

'malo šale nikada ne škodi
Private Sub Timer2_Timer()
    
    If Me.Visible = True Then
        num2 = num2 + 1
        If num2 > 500 Then num2 = num1 = 0
        
        If mX <> Me.left Or mY <> Me.top Then
            
            num1 = num1 + 1
            If num1 > 100 Then
                MsgBox "Well now... I can't fly there whole time!", vbQuestion
                num1 = 0
            End If
            
        End If
        
        mX = Me.left
        mY = Me.top
    End If
    
End Sub

'promjena vrijednosti varijable
Private Sub txtValue_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        
        If Me.lstVars.ListIndex >= 0 Then
            frmMain.saveVar Me.lstVars.List(Me.lstVars.ListIndex), Me.txtValue.Text
            frmMain.callCalc
        End If
        
    End If
    
End Sub
