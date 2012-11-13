VERSION 5.00
Begin VB.Form frmOpen 
   Caption         =   "Welcome to The Quest for The Cup"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   Picture         =   "frmOpen.frx":0000
   ScaleHeight     =   5700
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   3960
      TabIndex        =   1
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter The NHL Draft"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   1335
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Quest for The Cup~Minnesota Wild Trivia Game

'Author: Chris Adams

'Date: November 2007

'This program is a Minnesota Wild Trivia game that puts the user in the position of an NHL player.
'Like most players, users will start with the minor league team and work their way up to elite status.


Private Sub cmdEnter_Click()

    'Have the user enter their first amd last name for entry into the NHL Draft through an Input box
    PlayerFirst = InputBox("Welcome, Please enter your first name.", "First Name Please")
    PlayerLast = InputBox("Please enter your last name.", "Last Name Please")
    frmPos.Show
    frmOpen.Hide

End Sub

Private Sub cmdQuit_Click()

    'Show form Sources
    frmOpen.Hide
    frmSources.Show

End Sub
