VERSION 5.00
Begin VB.Form RachelHaney4 
   BackColor       =   &H0000FFFF&
   Caption         =   "RachelHaney4"
   ClientHeight    =   4875
   ClientLeft      =   3255
   ClientTop       =   2565
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   6435
   Visible         =   0   'False
   Begin VB.PictureBox picResults 
      Height          =   1095
      Left            =   360
      ScaleHeight     =   1035
      ScaleWidth      =   4275
      TabIndex        =   9
      Top             =   3480
      Width           =   4335
   End
   Begin VB.PictureBox picSuper 
      BackColor       =   &H0000FFFF&
      Height          =   735
      Left            =   3960
      Picture         =   "RachelHaneyVBProject4.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   1275
      TabIndex        =   8
      Top             =   2040
      Width           =   1335
   End
   Begin VB.PictureBox picMarriott 
      BackColor       =   &H0000FFFF&
      Height          =   975
      Left            =   2280
      Picture         =   "RachelHaneyVBProject4.frx":3342
      ScaleHeight     =   915
      ScaleWidth      =   1275
      TabIndex        =   7
      Top             =   2040
      Width           =   1335
   End
   Begin VB.PictureBox picCondo 
      BackColor       =   &H0000FFFF&
      Height          =   975
      Left            =   600
      Picture         =   "RachelHaneyVBProject4.frx":7894
      ScaleHeight     =   915
      ScaleWidth      =   1275
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   5160
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   615
      Left            =   4920
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdSuper 
      Caption         =   "Super 8 Hotel for $250.00"
      Height          =   855
      Left            =   3960
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdMarriott 
      Caption         =   "Marriott Hotel for $1500.000"
      Height          =   855
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdCondo 
      Caption         =   "A Condominium for $5000.00"
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblStay 
      BackColor       =   &H00FF80FF&
      Caption         =   "Where would you like to stay?"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "RachelHaney4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'RachelHaney4 (RachelHaneyVBProject3.frm)
'Rachel Haney 3/11/04
'This form asks people where they would like to stay
'during their vacation.

Private Sub cmdCondo_Click()
    Room = 1
    Total = Total + 5000
    picResults.Print "You decided to stay in a condo during your vacation."
    cmdContinue.Visible = True
    cmdMarriott.Visible = False
    cmdSuper.Visible = False
    cmdCondo.Visible = False
End Sub

Private Sub cmdContinue_Click()
    RachelHaney4.Visible = False
    RachelHaney5.Visible = True
    RachelHaney5.cmdContinue.Visible = False
End Sub

Private Sub cmdMarriott_Click()
    Room = 2
    Total = Total + 1500
    picResults.Print "You decided to stay in the luxerious Marriott Hotel"
    picResults.Print "during your vacation."
    cmdContinue.Visible = True
    cmdCondo.Visible = False
    cmdSuper.Visible = False
    cmdMarriott.Visible = False
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSuper_Click()
    Room = 3
    Total = Total + 250
    picResults.Print "You decided to stay in the Super 8 during your vacation."
    cmdContinue.Visible = True
    cmdCondo.Visible = False
    cmdMarriott.Visible = False
    cmdSuper.Visible = False
End Sub

