VERSION 5.00
Begin VB.Form RachelHaney2 
   BackColor       =   &H0000FFFF&
   Caption         =   "RachelHaney2"
   ClientHeight    =   4185
   ClientLeft      =   2955
   ClientTop       =   2880
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   6420
   Begin VB.PictureBox picResults 
      Height          =   615
      Left            =   2880
      ScaleHeight     =   555
      ScaleWidth      =   3315
      TabIndex        =   11
      Top             =   3360
      Width           =   3375
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H0000FFFF&
      Height          =   855
      Left            =   4680
      Picture         =   "RachelHaneyVBProject2.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FFFF&
      Height          =   1575
      Left            =   480
      Picture         =   "RachelHaneyVBProject2.frx":4112
      ScaleHeight     =   1515
      ScaleWidth      =   1035
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0000FFFF&
      Height          =   1455
      Left            =   3360
      Picture         =   "RachelHaneyVBProject2.frx":9848
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   8
      Top             =   720
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FFFF&
      Height          =   855
      Left            =   1920
      Picture         =   "RachelHaneyVBProject2.frx":E58E
      ScaleHeight     =   795
      ScaleWidth      =   1035
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdLondon 
      Caption         =   "London"
      Height          =   615
      Left            =   4920
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdParis 
      Caption         =   "Paris"
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdMiami 
      Caption         =   "Miami"
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdNYC 
      Caption         =   "New York City"
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblVisit 
      BackColor       =   &H00FF80FF&
      Caption         =   "Where would you like to visit?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "RachelHaney2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'RachelHaney2 (RachelHaneyVBProject1.frm)
'Rachel Haney 3/11/04
'This form asks the people where they would like
'to take their vacation.

Private Sub cmdContinue_Click()
'the code under the continue botton on this form and all
'of the following forms allows the user to only see
'one form at a time by making the current form invisible
'and the next form visible
    RachelHaney2.Visible = False
    RachelHaney3.Visible = True
    RachelHaney3.cmdContinue.Visible = False
    RachelHaney3.Show
'Display transportation in decending order
    picResults.Print "Transportation"; Tab(20); "Price"
    For J = 1 To CTR
        RachelHaney3.picResults.Print Vehicle(J); Tab(20); FormatCurrency(Price(J))
    Next J
    Total = 0
End Sub

Private Sub cmdLondon_Click()
'all of the buttons with the names of cities allow the user
'to decide where they will take their vacation.  Once a button
'is pushed the other buttons become invisible so the user
'can only choose one option.  The same is true for the choices
'on all of the following forms.
    City = 4
    picResults.Print "You chose London as your destination."
    cmdContinue.Visible = True
    cmdNYC.Visible = False
    cmdMiami.Visible = False
    cmdParis.Visible = False
    cmdLondon.Visible = False
End Sub

Private Sub cmdMiami_Click()
    City = 2
    picResults.Print "You chose Miami as your destination."
    cmdContinue.Visible = True
    cmdNYC.Visible = False
    cmdParis.Visible = False
    cmdLondon.Visible = False
    cmdMiami.Visible = False
End Sub

Private Sub cmdNYC_Click()
    City = 1
    picResults.Print "You chose New York City as your destination."
    cmdContinue.Visible = True
    cmdParis.Visible = False
    cmdMiami.Visible = False
    cmdLondon.Visible = False
    cmdNYC.Visible = False
End Sub

Private Sub cmdParis_Click()
    City = 3
    picResults.Print "You chose Paris as your destination."
    cmdContinue.Visible = True
    cmdNYC.Visible = False
    cmdMiami.Visible = False
    cmdLondon.Visible = False
    cmdParis.Visible = False
End Sub

Private Sub cmdQuit_Click()
    End
End Sub


