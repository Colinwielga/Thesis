VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H80000009&
   Caption         =   "IntroPage"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   Picture         =   "frmIntro.frx":0000
   ScaleHeight     =   5805
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label lblMike 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By Mike Patnode, 2006"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   5160
      Width           =   7455
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Minnesota Twins 2006"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
    frmHittingStats.Show
    frmPitchingStats.Hide
    frmCalcOwnStats.Hide
    frmSortTwins.Hide
    frmIntro.Hide
End Sub
