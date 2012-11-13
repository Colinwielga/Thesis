VERSION 5.00
Begin VB.Form frmHome 
   BackColor       =   &H000000FF&
   Caption         =   "Home"
   ClientHeight    =   6045
   ClientLeft      =   795
   ClientTop       =   2895
   ClientWidth     =   10275
   FillColor       =   &H000000FF&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   10275
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   7560
      TabIndex        =   9
      Top             =   4800
      Width           =   1455
   End
   Begin VB.PictureBox picClemens 
      Height          =   3255
      Left            =   5040
      Picture         =   "frmHome.frx":0000
      ScaleHeight     =   3195
      ScaleWidth      =   5235
      TabIndex        =   8
      Top             =   1320
      Width           =   5295
   End
   Begin VB.CommandButton cmdCredits 
      Caption         =   "Credits"
      Height          =   735
      Left            =   6120
      TabIndex        =   5
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuiz 
      Caption         =   "Quiz"
      Height          =   735
      Left            =   4800
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdStatistics 
      Caption         =   "Statistics"
      Height          =   735
      Left            =   3360
      TabIndex        =   3
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdSchedule 
      Caption         =   "Schedule"
      Height          =   735
      Left            =   1920
      TabIndex        =   2
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdStarters 
      Caption         =   "Starters"
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   4800
      Width           =   1335
   End
   Begin VB.PictureBox picGags 
      Height          =   4575
      Left            =   0
      Picture         =   "frmHome.frx":2AF32
      ScaleHeight     =   4515
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   2895
         Left            =   5040
         TabIndex        =   6
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.Label lblsju 
      BackColor       =   &H000000FF&
      Caption         =   "St. John's University (MN) Football"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5040
      TabIndex        =   7
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Each of these Private Subs Moves the user to the corresponding form

Private Sub cmdCredits_Click()
frmHome.Hide
frmCredits.Show
End Sub

Private Sub cmdQuit_Click()
 End
End Sub

Private Sub cmdQuiz_Click()
    frmHome.Hide
    frmQuiz.Show
End Sub

Private Sub cmdSchedule_Click()
    frmHome.Hide
    frmSchedule.Show
End Sub

Private Sub cmdStarters_Click()
 frmHome.Hide
 frmRoster.Show
End Sub

Private Sub cmdStatistics_Click()
frmHome.Hide
frmStatsHome.Show
End Sub
