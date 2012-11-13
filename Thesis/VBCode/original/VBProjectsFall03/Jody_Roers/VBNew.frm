VERSION 5.00
Begin VB.Form Comm 
   BackColor       =   &H80000007&
   Caption         =   "Comm"
   ClientHeight    =   9105
   ClientLeft      =   675
   ClientTop       =   855
   ClientWidth     =   13860
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   13860
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Return to Main Menu"
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin VB.PictureBox picComm 
      Height          =   1935
      Left            =   600
      Picture         =   "VBNew.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   975
      Left            =   9960
      TabIndex        =   3
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "What was the balance?"
      Height          =   975
      Left            =   6600
      TabIndex        =   2
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Can the Communication Club cover the expense?"
      Height          =   975
      Left            =   3480
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      Height          =   2295
      Left            =   3120
      ScaleHeight     =   2235
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   840
      Width           =   7935
   End
   Begin VB.Label lblName 
      Caption         =   "Created by Jody Roers"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblComm 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Communication Club"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Comm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Club Auditing Aid (VBProject.vbp)
'Form Name: Comm (VBNew.frm)
'Author: Jody Roers
'Date Written: 27 October 2003
'Purpose: to aid in my responsibilities as a Club Auditor in CSB Senate.
'   The program can give me the balance for any date in October selected and
'   determine whether the club has enough money for a specificied withdrawal.

Private Sub cmdLoad_Click()
Dim D As Integer
Dim M As Single
Dim Balance(1 To 31) As Single
Dim Sum As Single
Dim J As Integer
picResults.Cls 'clear screen
D = InputBox("What is the October Date? Only enter the number of the day specified.", "Enter Date") 'ask for october date user wants to withdraw money on
Do While D > 31 Or D < 1
    'if number isn't between 1 and 31(31 days in October) then give error message
    MsgBox "Sorry, you have entered an invalid date", , "Error"
    D = InputBox("What is the October Date?", "Enter Date") 'ask october date again
Loop 'continue until number entered is between 1 and 31
Open Menu.Path & "VB 10-21-02\Commtxt.txt" For Input As #1 'open file
For J = 1 To 31
    Input #1, Balance(J) 'fill it array
Next J
M = InputBox("Subtract how much money?", "Money") 'ask for amount of money to subtract
If Balance(D) > M Then 'checking if comm club has enough money to subtract that withdrawal
        Sum = Balance(D) - M 'club has enough money & is finding new balance in light of withdrawal
        picResults.Print "The Communication Club has enough money to cover this. The new balance is "; FormatCurrency(Sum)
        'printing new balance
    Else
        picResults.Print "The Communication Club does not have enough money to cover this expense."
        'letting user know that comm club doesn't have enough money
End If
Close #1 'close file
End Sub

Private Sub cmdMenu_Click()
Comm.Hide 'go back to main menu
Menu.Show

End Sub

Private Sub cmdQuit_Click()
End 'quit program
End Sub

Private Sub cmdSearch_Click()
Dim D As Integer
Dim J As Integer
Dim Balance(1 To 31) As Single
picResults.Cls 'clear screen
D = InputBox("What is the October Date?", "Enter Date") 'ask for october date user wants the balance on
Do While D > 31 Or D < 1
    'if user enters number less than 1 or more than 31 they get an error message
    MsgBox "Sorry, you have entered an invalid date", , "Error"
    D = InputBox("What is the October Date?", "Enter Date") 'ask october date again
Loop 'continue until user enters number between 1 and 31
Open Menu.Path & "VB 10-21-02\Commtxt.txt" For Input As #1 'open file with balances in it
For J = 1 To 31 'fill in array
    Input #1, Balance(J)
Next J
picResults.Print "The balance on October"; D; "was "; FormatCurrency(Balance(D)); "."
'print the balance on the day the user requested
Close #1 'close file
End Sub

