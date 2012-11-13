VERSION 5.00
Begin VB.Form frmMainMenu 
   Caption         =   "MainMenu"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   Picture         =   "frmMainMenu.frx":0000
   ScaleHeight     =   9060
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "For More Information on Jeopardy... "
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4320
      TabIndex        =   6
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdWinners 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Winner winner chicken dinner..."
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5280
      TabIndex        =   5
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit Game "
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1920
      TabIndex        =   3
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdCitations 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Works Cited"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3000
      TabIndex        =   2
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play Game!!!"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1920
      TabIndex        =   1
      Top             =   3360
      Width           =   4215
   End
   Begin VB.CommandButton cmdRules 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rules of the Game"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   600
      TabIndex        =   0
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label lblCSB 
      BackColor       =   &H80000013&
      Caption         =   "CSB/SJU EDITION"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   4
      Top             =   2280
      Width           =   6015
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Our project name is: CSB/SJU Jeopardy
'Form name: frmMainMenu
'Authors: Emma Jaynes, Lindsay Havlik, Brooke Beyer
'Date Written: 10/26/08
'Objective: This project asks the user trivia questions about CSB/SJU and awards points/dollars for correct answers "Jeopardy style".
'Comments: Each button takes the user to another form.  1.Rules  2.Works Cited  3.End the game  4.Play the game

Private Sub cmdCitations_Click()
frmWorksCited.Show 'shows the works cited page
frmMainMenu.Hide 'hides the main menu form

End Sub

Private Sub cmdPlay_Click()

frmGameBoard.Show
frmMainMenu.Hide

Contestant = InputBox("Welcome to Jeopardy! Please enter your name:")
MsgBox ("Welcome " & Contestant & " Let's Play Jeopardy!")

frmGameBoard.cmdDouble.Enabled = False

frmGameBoard.picContestant.Print Contestant

End Sub

Private Sub cmdQuit_Click()
End ' quits the game

End Sub

Private Sub cmdRules_Click()
frmMainMenu.Hide 'hides the main menu form
frmRules.Show 'shows the rules form

End Sub

Private Sub Label1_Click()

End Sub

Private Sub cmdWinners_Click()
frmMainMenu.Hide
frmWinners.Show

End Sub

Private Sub Command1_Click()
frmMainMenu.Hide
frmInfo.Show

End Sub
