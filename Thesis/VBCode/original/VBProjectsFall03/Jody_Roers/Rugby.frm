VERSION 5.00
Begin VB.Form Rugby 
   BackColor       =   &H80000007&
   Caption         =   "Rugby"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   12360
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   10200
      TabIndex        =   6
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "What was the balance?"
      Height          =   1095
      Left            =   7320
      TabIndex        =   5
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Can the Women's Rugby Club cover the expense?"
      Height          =   1095
      Left            =   4440
      TabIndex        =   4
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Return to the Main Menu"
      Height          =   1095
      Left            =   1080
      TabIndex        =   3
      Top             =   4560
      Width           =   1455
   End
   Begin VB.PictureBox picRugby 
      Height          =   3375
      Left            =   240
      Picture         =   "Rugby.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   4155
      TabIndex        =   2
      Top             =   600
      Width           =   4215
   End
   Begin VB.PictureBox picResults 
      Height          =   2055
      Left            =   5040
      ScaleHeight     =   1995
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   1560
      Width           =   6495
   End
   Begin VB.Label lblName 
      Caption         =   "Created by Jody Roers"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lbRugby 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Women's Rugby Club"
      BeginProperty Font 
         Name            =   "Modern"
         Size            =   26.25
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   1
      Top             =   600
      Width           =   4455
   End
End
Attribute VB_Name = "Rugby"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Club Auditing Aid (VBProject.vbp)
'Form Name: Rugby (Rugby.frm)
'Author: Jody Roers
'Date Written: 27 October 2003
'Purpose: to aid in my responsibilities as a Club Auditor in CSB Senate.
'   The program can give me the balance for any date in October selected and
'   determine whether the club has enough money for a specificied withdrawal.

Private Sub cmdLoad_Click()
Dim D As Integer
Dim M As Single
Dim J As Integer
Dim Balance(1 To 31) As Single
Dim Sum As Single
picResults.Cls
D = InputBox("What is the October Date?", "Enter Date") 'ask for october date user wants to withdraw money on
Do While D > 31 Or D < 1 'if number isn't between 1 and 31(31 days in October) then give error message
    MsgBox "Sorry, you have entered an invalid date", , "Error"
    D = InputBox("What is the October Date? Only enter the number of the day specified.", "Enter Date") 'ask october date again
Loop 'continue until number entered is between 1 and 31
Open Menu.Path & "VB 10-21-02\Rugbytxt.txt" For Input As #1 'open file
For J = 1 To 31 'fill array
    Input #1, Balance(J)
Next J
M = InputBox("Subtract how much money?", "Money") 'ask for amount of money to subtract
If Balance(D) > M Then 'checking if club has enough money to subtract that withdrawal
        Sum = Balance(D) - M 'club has enough money and finding new balance in light of withdrawal
        picResults.Print "The Women's Rugby Club has enough money to cover this. The new balance is "; FormatCurrency(Sum)
        'printing new balance
    Else
        picResults.Print "The Women's Rugby Club does not have enough money to cover this expense."
        'letting user know that comm club doesn't have enough money
End If
Close #1 'close file
End Sub


Private Sub cmdMenu_Click()
Rugby.Hide 'return to main menu
Menu.Show

End Sub

Private Sub cmdQuit_Click()
End 'quit
End Sub

Private Sub cmdSearch_Click()
Dim D As Integer
Dim J As Integer
Dim Balance(1 To 31) As Single
picResults.Cls 'clear screen
D = InputBox("What is the October Date?", "Enter Date") 'ask for october date user wants the balance on
Do While D > 31 Or D < 1 'if user enters number less than 1 or more than 31 they get an error message
    MsgBox "Sorry, you have entered an invalid date", , "Error"
    D = InputBox("What is the October Date?", "Enter Date") 'ask october date again
Loop 'continue until user enters number between 1 and 31
Open Menu.Path & "VB 10-21-02\Rugbytxt.txt" For Input As #1 'open file
For J = 1 To 31 'file array
    Input #1, Balance(J)
Next J
picResults.Print "The balance on October"; D; "was "; FormatCurrency(Balance(D)); "."
'print the balance on the day the user requested
Close #1 'close file
End Sub



