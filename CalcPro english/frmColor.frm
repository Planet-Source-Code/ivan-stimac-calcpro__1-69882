VERSION 5.00
Begin VB.Form frmColor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select graph color"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   47
      Left            =   3000
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   39
      Top             =   1920
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   46
      Left            =   2580
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   38
      Top             =   1920
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   45
      Left            =   2160
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   37
      Top             =   1920
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   44
      Left            =   1740
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   36
      Top             =   1920
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00004040&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   43
      Left            =   1320
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   35
      Top             =   1920
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   42
      Left            =   900
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   34
      Top             =   1920
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   41
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   33
      Top             =   1920
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   40
      Left            =   60
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   32
      Top             =   1920
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00800080&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   37
      Left            =   3000
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   31
      Top             =   1500
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   36
      Left            =   2580
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   30
      Top             =   1500
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   35
      Left            =   2160
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   29
      Top             =   1500
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   34
      Left            =   1740
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   28
      Top             =   1500
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   33
      Left            =   1320
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   27
      Top             =   1500
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   32
      Left            =   900
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   26
      Top             =   1500
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   31
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   25
      Top             =   1500
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   30
      Left            =   60
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   24
      Top             =   1500
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   27
      Left            =   3000
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   23
      Top             =   1080
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   26
      Left            =   2580
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   22
      Top             =   1080
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   25
      Left            =   2160
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   21
      Top             =   1080
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   24
      Left            =   1740
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   20
      Top             =   1080
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   23
      Left            =   1320
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   19
      Top             =   1080
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   22
      Left            =   900
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   18
      Top             =   1080
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   21
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   17
      Top             =   1080
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   20
      Left            =   60
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   16
      Top             =   1080
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   17
      Left            =   3000
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   15
      Top             =   660
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   16
      Left            =   2580
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   14
      Top             =   660
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   15
      Left            =   2160
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   13
      Top             =   660
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   14
      Left            =   1740
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   12
      Top             =   660
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   13
      Left            =   1320
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   11
      Top             =   660
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   12
      Left            =   900
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   10
      Top             =   660
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   11
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   9
      Top             =   660
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   10
      Left            =   60
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   8
      Top             =   660
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   7
      Left            =   3000
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   7
      Top             =   240
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   6
      Left            =   2580
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   6
      Top             =   240
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   5
      Left            =   2160
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   5
      Top             =   240
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   4
      Left            =   1740
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   240
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   3
      Left            =   1320
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   3
      Top             =   240
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   2
      Left            =   900
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   2
      Top             =   240
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   1
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   1
      Top             =   240
      Width           =   400
   End
   Begin VB.PictureBox pcColSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   0
      Left            =   60
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   240
      Width           =   400
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub pcColSel_Click(Index As Integer)

    frmOpcije.pcBoja.BackColor = Me.pcColSel(Index).BackColor
    Me.Hide
    
End Sub
