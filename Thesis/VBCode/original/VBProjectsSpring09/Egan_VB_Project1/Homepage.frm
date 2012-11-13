VERSION 5.00
Begin VB.Form frmHomepage 
   BackColor       =   &H0000C000&
   Caption         =   "Welcome"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Exit"
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdGolf 
      BackColor       =   &H0000FFFF&
      Caption         =   "Golf"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdNHL 
      BackColor       =   &H0000FFFF&
      Caption         =   "NHL"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdNCAAB 
      BackColor       =   &H0000FFFF&
      Caption         =   "NCAAB"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdNCAAF 
      BackColor       =   &H0000FFFF&
      Caption         =   "NCAAF"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdSoccer 
      BackColor       =   &H0000FFFF&
      Caption         =   "Soccer"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdMLB 
      BackColor       =   &H0000FFFF&
      Caption         =   "MLB"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdNBA 
      BackColor       =   &H0000FFFF&
      Caption         =   "NBA"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdNFL 
      BackColor       =   &H0000FFFF&
      Caption         =   "NFL"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblNote3 
      BackColor       =   &H0000C000&
      Caption         =   "*Note: Only bet on one team at a time."
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   4920
      Width           =   3255
   End
   Begin VB.Label lblNote2 
      BackColor       =   &H0000C000&
      Caption         =   "*Note: You may bet on the same game more than once."
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   4560
      Width           =   3975
   End
   Begin VB.Label lblNote1 
      BackColor       =   &H0000C000&
      Caption         =   "*Note: All winning bets pay double."
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Label lblPickASport 
      BackColor       =   &H0000C000&
      Caption         =   "Which sport would you like to bet on?"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label lblPlaceYourBets 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Place Your Bets!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "frmHomepage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Sports Betting Project
'frmHomepage
'Written by: Sean Egan
'Written on: 3/15/09
'This form is the launching point where the user may choose
' a number of different sports to bet on.

Private Sub cmdExit_Click()
    'Ends the program
    End
End Sub

Private Sub cmdGolf_Click()
    'Loads the Golf form
    frmGolf.Show
    'Hides the Homepage
    frmHomepage.Hide
End Sub

Private Sub cmdMLB_Click()
    'Loads the MLB form
    frmMLB.Show
    'Hides the Homepage
    frmHomepage.Hide
End Sub

Private Sub cmdNBA_Click()
    'Loads the NBA form
    frmNBA.Show
    'Hides the Homepage
    frmHomepage.Hide
End Sub

Private Sub cmdNCAAB_Click()
    'Loads the NCAA Basketball form
    frmNCAAB.Show
    'Hides the Homepage
    frmHomepage.Hide
End Sub

Private Sub cmdNCAAF_Click()
    'Loads the NCAA Football form
    frmNCAAF.Show
    'Hides the Homepage
    frmHomepage.Hide
End Sub

Private Sub cmdNFL_Click()
    'Loads the NFL form
    frmNFL.Show
    'Hides the Homepage
    frmHomepage.Hide
End Sub

Private Sub cmdNHL_Click()
    'Loads the NHL form
    frmNHL.Show
    'Hides the Homepage
    frmHomepage.Hide
End Sub

Private Sub cmdSoccer_Click()
    'Loads the Soccer form
    frmSoccer.Show
    'Hides the Homepage
    frmHomepage.Hide
End Sub


