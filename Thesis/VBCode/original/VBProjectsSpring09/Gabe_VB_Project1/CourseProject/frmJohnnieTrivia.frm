VERSION 5.00
Begin VB.Form frmJohnnieTrivia 
   Caption         =   "Johnnie Trivia"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   Picture         =   "frmJohnnieTrivia.frx":0000
   ScaleHeight     =   7890
   ScaleWidth      =   11685
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Start the questions!"
      Height          =   975
      Left            =   8640
      TabIndex        =   1
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Back to Main Menu"
      Height          =   975
      Left            =   8760
      TabIndex        =   0
      Top             =   4200
      Width           =   2055
   End
End
Attribute VB_Name = "frmJohnnieTrivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Fun with CSB/SJU History!
'frmJohnnieTrivia
'Audrey Gabe
'Written 3/16/09
'Similar to the Bennie trivia section, this form allows the user to answer trivia questions about St. John's and see if they are right

Private Sub cmdMenu_Click()
frmJohnnieTrivia.Hide
frmMenu.Show
End Sub

Private Sub cmdPlay_Click()
Dim K As String

K = InputBox("Who was the first president of Saint John's University?", , "Enter answer here") 'User enters answer to a question
If K = "Abbot Rupert Seidenbusch" Then
    MsgBox "Great!", , "Your answer is correct!" 'If answer is correct
    Else
        MsgBox "Sorry, the answer is Abbot Rupert Seidenbusch", , "Incorrect" 'If answer is incorrect
End If

K = InputBox("On what building are the letters that represent 'That in all things God may be glorified?'", , "Enter answer here")
If K = "the Quad" Then
    MsgBox "You're doing great!", , "Your answer is correct!"
    Else
        MsgBox "Sorry, the answer is the Quad", , "Incorrect"
End If

K = InputBox("In what year was SJU established in its current location?", , "Enter answer here")
If K = "1866" Then
    MsgBox "Super!", , "Your answer is correct!"
    Else
        MsgBox "Sorry, the answer is 1866", , "Incorrect"
End If

frmJohnnieTrivia.Hide
frmMenu.Show


End Sub


