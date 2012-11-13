VERSION 5.00
Begin VB.Form RachelHaney5 
   BackColor       =   &H0000FFFF&
   Caption         =   "RachelHaney5"
   ClientHeight    =   4935
   ClientLeft      =   3255
   ClientTop       =   2265
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   6315
   Visible         =   0   'False
   Begin VB.PictureBox picResults 
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1275
      ScaleWidth      =   4875
      TabIndex        =   9
      Top             =   3480
      Width           =   4935
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0000FFFF&
      Height          =   855
      Left            =   720
      Picture         =   "RachelHaneyVBProject5.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   4440
      Picture         =   "RachelHaneyVBProject5.frx":3CD2
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FFFF&
      Height          =   1455
      Left            =   2640
      Picture         =   "RachelHaneyVBProject5.frx":9624
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   5160
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   615
      Left            =   5160
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdRelatives 
      Caption         =   "Visit the Relatives for $200.00"
      Height          =   735
      Left            =   2640
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdSight 
      Caption         =   "Sight see for $1000.00"
      Height          =   735
      Left            =   4560
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdBeach 
      Caption         =   "Visit the Beach  FREE!"
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblTravel 
      BackColor       =   &H00FF80FF&
      Caption         =   "What would you like to do on your vacation?"
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
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "RachelHaney5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'RachelHaney5 (RachelHaneyVBProject4.frm)
'Rachel Haney 3/11/04
'This form asks people what they would like to do on
'their vacation.

Private Sub cmdBeach_Click()
    Visit = 1
    picResults.Print "You decided to spend your time at the beach during your vacation."
    picResults.Print "Great choice!"
    cmdContinue.Visible = True
    cmdRelatives.Visible = False
    cmdSight.Visible = False
    cmdBeach.Visible = False
End Sub

Private Sub cmdContinue_Click()
    RachelHaney5.Visible = False
    RachelHaney6.Visible = True
    RachelHaney6.cmdContinue.Visible = False
    RachelHaney6.cmdTotal.Visible = False
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdRelatives_Click()
    Visit = 3
    Total = Total + 200
    picResults.Print "You decided to take a vacation and visit your relatives."
    picResults.Print "How sweet of you!"
    cmdContinue.Visible = True
    cmdBeach.Visible = False
    cmdSight.Visible = False
    cmdRelatives.Visible = False
End Sub

Private Sub cmdSight_Click()
    Visit = 2
    Total = Total + 1000
    picResults.Print "You decided to sight see during your vacation."
    cmdContinue.Visible = True
    cmdBeach.Visible = False
    cmdRelatives.Visible = False
    cmdSight.Visible = False
End Sub

