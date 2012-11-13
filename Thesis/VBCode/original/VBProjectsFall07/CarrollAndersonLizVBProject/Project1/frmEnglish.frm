VERSION 5.00
Begin VB.Form frmEnglish 
   Caption         =   "English"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   Picture         =   "frmEnglish.frx":0000
   ScaleHeight     =   5280
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBackToMain 
      Caption         =   "Back to Main Screen"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdHangMan 
      Caption         =   "Hang Man"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdUnscramble 
      Caption         =   "Unscramble Words"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   1920
      Width           =   1575
   End
End
Attribute VB_Name = "frmEnglish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBackToMain_Click()

'takes user back to the main screen

frmEnglish.Visible = False
frmMainScreen.Visible = True
End Sub

Private Sub cmdHangMan_Click()

'Opens Hangman application

frmEnglish.Visible = False
frmHangman.Visible = True
End Sub

Private Sub cmdQuit_Click()

'Tells user good luck with thier homework, and quits

MsgBox ("Good luck with your " & Homework & " hours of homework!")
End
End Sub

Private Sub cmdUnscramble_Click()

'opens unscramble word page

frmEnglish.Hide
frmUnscramble.Show
MsgBox ("The theme of the words to be unscrambled is states in the US")
End Sub
