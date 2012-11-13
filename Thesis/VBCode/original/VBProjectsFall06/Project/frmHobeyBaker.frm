VERSION 5.00
Begin VB.Form frmHobeyBaker 
   BackColor       =   &H00000080&
   Caption         =   "Hobey Baker"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   855
      Index           =   1
      Left            =   2640
      TabIndex        =   11
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   855
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton cmdStauber 
      BackColor       =   &H00C0C0C0&
      Height          =   2775
      Index           =   1
      Left            =   5400
      Picture         =   "frmHobeyBaker.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdBroten 
      BackColor       =   &H00C0C0C0&
      Height          =   2775
      Index           =   0
      Left            =   7920
      Picture         =   "frmHobeyBaker.frx":0BE4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdBonin 
      Height          =   2775
      Left            =   2880
      Picture         =   "frmHobeyBaker.frx":1FD8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdLeopold 
      Height          =   2775
      Left            =   360
      Picture         =   "frmHobeyBaker.frx":339A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label lblBonin 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Brian Bonin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   3
      Left            =   3480
      TabIndex        =   9
      Top             =   4800
      Width           =   1230
   End
   Begin VB.Label lbStauber 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Robb Stauber"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   2
      Left            =   5880
      TabIndex        =   8
      Top             =   4800
      Width           =   1500
   End
   Begin VB.Label lbBroten 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Neal Broten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      Left            =   8520
      TabIndex        =   7
      Top             =   4800
      Width           =   1260
   End
   Begin VB.Label lblLeopold 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Jordan Leopold"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   720
      TabIndex        =   6
      Top             =   4800
      Width           =   1650
   End
   Begin VB.Label blbClick 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Click on the picture below to view profile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3000
      TabIndex        =   5
      Top             =   1200
      Width           =   4170
   End
   Begin VB.Label lblHobey 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "Hobey Baker Award Winners"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   6195
   End
End
Attribute VB_Name = "frmHobeyBaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmHobeyBaker
'Cole and John
'10/30/06
'Objective: The objective of this form is to allow the user to view past Hobey Baker
'award winners.  The user can do this by clicking on the apprpriate player to view
'their information, statistics, and accomplishments.

Option Explicit
Private Sub cmdBack_Click(Index As Integer)
    frmHobeyBaker.Visible = False           'allows user to go back
    frmHistory.Visible = True
End Sub

Private Sub cmdBonin_Click()
    frmBonin.Visible = True                 'makes visible the Bonin form, hides Hobey Baker form
    frmHobeyBaker.Visible = False
End Sub

Private Sub cmdBroten_Click(Index As Integer)
    frmBroten.Visible = True
    frmHobeyBaker.Visible = False
End Sub

Private Sub cmdHome_Click(Index As Integer)
    frmHobeyBaker.Visible = False
    frmMain.Visible = True
End Sub

Private Sub cmdLeopold_Click()
    frmHobeyBaker.Visible = False
    frmLeopold.Visible = True
End Sub

Private Sub cmdStauber_Click(Index As Integer)
    frmStauber.Visible = True
    frmHobeyBaker.Visible = False
End Sub
