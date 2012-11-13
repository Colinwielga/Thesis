VERSION 5.00
Begin VB.Form frmPitchingPage
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   8205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn
      BackColor       =   &H0000FFFF&
      Caption         =   "Return to Main Page"
      BeginProperty Font
         Name            =   "Castellar"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton cmdCalcERA
      BackColor       =   &H0000FFFF&
      Caption         =   "Calculate ERA"
      BeginProperty Font
         Name            =   "Castellar"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox txtEnterEarnedRuns
      BackColor       =   &H00C0C0C0&
      BeginProperty Font
         Name            =   "Castellar"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   3
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox txtEnterInnings
      BackColor       =   &H00C0C0C0&
      BeginProperty Font
         Name            =   "Castellar"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   2
      Top             =   960
      Width           =   2655
   End
   Begin VB.Image Image1
      Height          =   1470
      Left            =   9000
      Picture         =   "PitchingPage.frx":0000
      Top             =   6240
      Width           =   1680
   End
   Begin VB.Label lblEnterEarnedRuns
      BackColor       =   &H0080FFFF&
      Caption         =   "Enter Earned Runs"
      BeginProperty Font
         Name            =   "Castellar"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   1
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label lblEnterIP
      BackColor       =   &H0080FFFF&
      Caption         =   "Enter Innings Pitched"
      BeginProperty Font
         Name            =   "Castellar"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "frmPitchingPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Calculate pitching stats - textboxes for user to easily calculate stats
'For pitchers based on game, seaons, career, etc.

Private Sub cmdCalcERA_Click()
'declaring variables
Dim InningsPitched As String

'defining variables
InningsPitched = txtEnterInnings.Text
Dim EarnedRuns As String
EarnedRuns = txtEnterEarnedRuns.Text
Dim ER As Single
ER = EarnedRuns / InningsPitched
Dim ERA As Single
ERA = ER * 7

'print to msgbox
MsgBox "Your pitcher's ERA is: " & FormatNumber(ERA, 3) & "."
End Sub

Private Sub Image1_Click()
MsgBox "Strike 'em out!"
End Sub

'return to main page
Private Sub cmdReturn_Click()
'clear textboxes

frmDeffensivePage.Hide
frmPitchingPage.Hide
frmOffensivePage.Hide
frmAHomePage.Show
End Sub
