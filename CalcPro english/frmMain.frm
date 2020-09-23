VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CalcPro"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7485
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox mFileLst 
      Height          =   300
      Left            =   4560
      Pattern         =   "*.inc"
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "x2000"
      Height          =   375
      Left            =   6300
      TabIndex        =   4
      Top             =   780
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtFunction 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1260
      Width           =   7095
   End
   Begin CalcPro.ProTabCTL mTab 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   3420
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   556
      StyleNormal     =   7
      StyleActive     =   7
      ScrollStyle     =   0
      TabAreaBackColor=   16777215
      ShadowColor     =   526344
      TabColorActive  =   16777215
      ForeColorActive =   8421504
      ForeColorHover  =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontActive {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabHeight       =   20
      TabSpacing      =   4
      FontSpacing     =   0
      ActiveTab       =   4
      EnableFontMoving=   0   'False
      Appearance      =   0
      DrawClientArea  =   0   'False
      TabCount        =   5
      TabCaption0     =   "Function table"
      TabCaption1     =   "Function graph"
      TabCaption2     =   "Variables"
      TabCaption3     =   "About"
      TabCaption4     =   "x"
      containdecControlsCount0=   0
      containdecControlsCount1=   0
      containdecControlsCount2=   0
      containdecControlsCount3=   0
      containdecControlsCount4=   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Input function:"
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   0
      Picture         =   "frmMain.frx":030A
      Top             =   0
      Width           =   7500
   End
   Begin VB.Label lblRes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   420
      TabIndex        =   2
      Top             =   2820
      Width           =   7815
   End
   Begin VB.Menu mnuDatoteka 
      Caption         =   "File"
      Begin VB.Menu buttClear 
         Caption         =   "Clear"
         Shortcut        =   ^D
      End
      Begin VB.Menu c1_mnuDatoteka 
         Caption         =   "-"
      End
      Begin VB.Menu buttIzlaz 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "Insert"
      Begin VB.Menu buttMnuInsF 
         Caption         =   "Function"
         Begin VB.Menu buttFunct 
            Caption         =   "#"
            Index           =   0
         End
      End
      Begin VB.Menu buttMnuLibF 
         Caption         =   "Funtion from library"
         Begin VB.Menu buttLibFunc 
            Caption         =   "#"
            Index           =   0
         End
      End
      Begin VB.Menu buttMnuOpIns 
         Caption         =   "Operator"
         Begin VB.Menu buttMnuAritm 
            Caption         =   "Arithmetic operators"
            Begin VB.Menu buttAritmetickiOp 
               Caption         =   " + "
               Index           =   0
            End
            Begin VB.Menu buttAritmetickiOp 
               Caption         =   " - "
               Index           =   1
            End
            Begin VB.Menu buttAritmetickiOp 
               Caption         =   " / "
               Index           =   2
            End
            Begin VB.Menu buttAritmetickiOp 
               Caption         =   " * "
               Index           =   3
            End
            Begin VB.Menu buttAritmetickiOp 
               Caption         =   " ^ "
               Index           =   4
            End
         End
         Begin VB.Menu buttMnuLogOp 
            Caption         =   "Logican operators and comparison operators"
            Begin VB.Menu buttLogicki 
               Caption         =   " == "
               Index           =   0
            End
            Begin VB.Menu buttLogicki 
               Caption         =   " !="
               Index           =   1
            End
            Begin VB.Menu buttLogicki 
               Caption         =   " <= "
               Index           =   2
            End
            Begin VB.Menu buttLogicki 
               Caption         =   " >="
               Index           =   3
            End
            Begin VB.Menu buttLogicki 
               Caption         =   " <"
               Index           =   4
            End
            Begin VB.Menu buttLogicki 
               Caption         =   " >"
               Index           =   5
            End
            Begin VB.Menu buttLogicki 
               Caption         =   " && "
               Index           =   6
            End
            Begin VB.Menu buttLogicki 
               Caption         =   " ||"
               Index           =   7
            End
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cCalc As New clsCalculator

Private Sub buttAritmetickiOp_Click(Index As Integer)

    Me.txtFunction.SelText = Me.buttAritmetickiOp(Index).Caption
    
End Sub

Private Sub buttClear_Click()

    Me.txtFunction.Text = vbNullString
    
End Sub

Private Sub buttFunct_Click(Index As Integer)

    Me.txtFunction.SelText = Me.buttFunct(Index).Caption
    
End Sub

Private Sub buttIzlaz_Click()

    Unload Me
    
End Sub

Private Sub buttLibFunc_Click(Index As Integer)
    Me.txtFunction.SelText = Me.buttLibFunc(Index).Caption
End Sub


Private Sub buttLogicki_Click(Index As Integer)

    Me.txtFunction.SelText = Me.buttLogicki(Index).Caption
    
End Sub


Private Sub Command1_Click()
    Dim i As Integer, mTmr As Double
    mTmr = Timer
    For i = 1 To 2000
        cCalc.Calculate Me.txtFunction.Text
    Next i
    MsgBox Timer - mTmr
    cCalc.ClearParseTree
End Sub

Private Sub Form_Load()
    
    Dim i As Integer, k As Integer
    Load frmOpcije
    Load frmColor
    frmOpcije.Width = Me.Width
    Me.mFileLst.Path = App.Path & "\include"
    
    'ucitavanje biblioteka
    For i = 0 To Me.mFileLst.ListCount - 1
        
        k = InStrRev(Me.mFileLst.List(i), ".")
        If LCase$(Mid$(Me.mFileLst.List(i), k + 1)) = "inc" Then
            cCalc.IncludeLib App.Path & "\include\" & Me.mFileLst.List(i)
        End If
        
    Next i
    '
    cCalc.LoadVarLib App.Path & "\include\predefvars.incvar"
    
    'popis glavnih funkcija
    For i = 1 To cCalc.FunctsCount
        
        If i > 1 Then
            Load Me.buttFunct(i - 1)
            Me.buttFunct(i - 1).Visible = True
        End If
        Me.buttFunct(i - 1).Caption = cCalc.FunctionName(i)
        
    Next i
    
    'popis funkcija iz biblioteke
    For i = 1 To cCalc.FunctsLibCount
        
        If i > 1 Then
            Load Me.buttLibFunc(i - 1)
            Me.buttLibFunc(i - 1).Visible = True
        End If
        Me.buttLibFunc(i - 1).Caption = cCalc.FunctionLibName(i)
        
    Next i

End Sub

'ucitavanje varijabli u listu
'   na frmopcije
Public Sub loadVars()

    Dim i As Integer
    frmOpcije.lstVars.Clear
    
    For i = 1 To cCalc.VarCount
        
        frmOpcije.lstVars.AddItem cCalc.varName(i)
        
    Next i
    
End Sub

'definiranje nove varijable
Public Sub saveVar(ByVal varName As String, ByVal varVal As String)
    
    cCalc.ErrorClear
    cCalc.defineVariable varName, varVal
    
    If cCalc.ErrorGetNumber <> 0 Then
        
        MsgBox "Invalid variable name or invalid value!", vbExclamation
    
    End If
    
End Sub

'poziv racunanja
Public Sub callCalc()

    txtFunction_Change
    
End Sub


Public Function getVarValue(ByVal varName As String) As String
    
    getVarValue = cCalc.VarValue(varName)
    
End Function



Private Sub Form_Unload(Cancel As Integer)
    
    Unload frmOpcije
    Unload frmColor
    End
    
End Sub


Private Sub mTab_Click()
    
    frmOpcije.frmVarijable.Visible = False
    frmOpcije.frmCrtanje.Visible = False
    frmOpcije.frmRacun.Visible = False
    frmOpcije.frmAbout.Visible = False
    '
    Select Case mTab.ActiveTab
        Case 0
            frmOpcije.frmRacun.Visible = True
            frmOpcije.Show
            
        Case 1
            frmOpcije.frmCrtanje.Visible = True
            frmOpcije.Show
            
        Case 2
            Me.loadVars
            frmOpcije.frmVarijable.Visible = True
            frmOpcije.Show
            
        Case 3
            frmOpcije.frmAbout.Visible = True
            frmOpcije.Show
            
        Case 4
            frmOpcije.Hide
            
    End Select
    
End Sub

Public Function getValue(ByVal strFunction As String, ByVal mX As Single) As Double
    
    cCalc.ErrorClear
    cCalc.defineVariable "$x", Str$(mX)
    getValue = Val(cCalc.Calculate(strFunction))
    '
    cCalc.ClearParseTree
    
End Function

Public Function getLastErr() As Long
    
    getLastErr = cCalc.ErrorGetNumber
    
End Function
Public Sub clearError()
    
    cCalc.ErrorClear
    
End Sub

Private Sub txtFunction_Change()
    
    cCalc.ErrorClear
    lblRes.Caption = " = " & cCalc.Calculate(Me.txtFunction.Text)
    
    If cCalc.ErrorGetNumber <> 0 Then
        
        lblRes.Caption = cCalc.ErrorGetDescription
        
    End If
    
    cCalc.ClearParseTree
    
End Sub

