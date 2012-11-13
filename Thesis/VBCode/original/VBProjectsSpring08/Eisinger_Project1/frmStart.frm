VERSION 5.00
Begin VB.Form frmStart 
   Caption         =   "Start"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   Picture         =   "frmStart.frx":0000
   ScaleHeight     =   5910
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNo 
      BackColor       =   &H00004080&
      Caption         =   "No"
      Height          =   1095
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdYes 
      BackColor       =   &H00004080&
      Caption         =   "Yes"
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004080&
      Caption         =   "Are you ready to Test your Baseball Knowledge?"
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1800
      Width           =   3735
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form either starts the program or ends the program
'It opens a text of questions and their values and turns that text
'into two arrays of questions and money.
'The questions are printed in their corresponding round forms in the
'picResults picture boxes in those forms.
'Askes the user to input their name into an input box
'Displays a welcome message in a message box
'Hides this form and shows the round 1 form


Private Sub cmdNo_Click()
'End the program
End
End Sub

Private Sub cmdYes_Click()
'open the file questionsandanswers.txt
'turn this file into two arrays called questions and money
Found = False
Round = 1
Open App.Path & "\baseball.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Questions(CTR), Money(CTR)
Loop

'Close the file
Close #1

'Use an input box to recieve the users name
'Print in a message box the players name along with a welcome message
'Hide this form and show the round1 form
frmStart.Hide
frmRound1.Show
PlayerName = InputBox("Enter your name", "Name")
MsgBox "Batter Up " & PlayerName & " to Baseball Knowledge!", , "Batter Up"
frmRound1.picResults.Print Questions(Round)
frmRound1.picMoney.Print FormatCurrency(Money(Round))
End Sub
