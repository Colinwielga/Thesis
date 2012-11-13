VERSION 5.00
Begin VB.Form Spanish 
   BackColor       =   &H80000007&
   Caption         =   "Spanish"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13740
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   13740
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picSpanish 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   480
      Picture         =   "VBSpanish.frx":0000
      ScaleHeight     =   1935
      ScaleWidth      =   2535
      TabIndex        =   6
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   9600
      TabIndex        =   4
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "What was the balance?"
      Height          =   1095
      Left            =   6600
      TabIndex        =   3
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Can the Spanish Club cover the expense?"
      Height          =   1095
      Left            =   3240
      TabIndex        =   2
      Top             =   4800
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      Height          =   2175
      Left            =   3720
      ScaleHeight     =   2115
      ScaleWidth      =   6915
      TabIndex        =   1
      Top             =   1680
      Width           =   6975
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Return to the Main Menu"
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label lblName 
      Caption         =   "Created by Jody Roers"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblSpanish 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Spanish Club"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "Spanish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Club Auditing Aid (VBProject.vbp)
'Form Name: Spanish (VBSpanish.frm)
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
Do While D > 31 Or D < 1 'if number isn't between 1 and 31(31 days in October) then give error message
    MsgBox "Sorry, you have entered an invalid date", , "Error"
    D = InputBox("What is the October Date?", "Enter Date") 'ask october date again
Loop 'continue until number entered is between 1 and 31
Open Menu.Path & "VB 10-21-02\Spanishtxt.txt" For Input As #1 'open file
For J = 1 To 31 'fill array
    Input #1, Balance(J)
Next J
M = InputBox("Subtract how much money?", "Money") 'ask for amount of money to subtract
If Balance(D) > M Then 'checking if club has enough money to subtract that withdrawal
        Sum = Balance(D) - M 'club has enough money and finding new balance in light of withdrawal
        picResults.Print "The Spanish Club has enough money to cover this. The new balance is "; FormatCurrency(Sum)
        'printing new balance
    Else
        picResults.Print "The Spanish Club does not have enough money to cover this expense."
        'letting user know that comm club doesn't have enough money
End If
Close #1 'close file
End Sub

Private Sub cmdMenu_Click()
Spanish.Hide 'return to main menu
Menu.Show
End Sub

Private Sub cmdQuit_Click()
End 'Quit
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
Open Menu.Path & "VB 10-21-02\Spanishtxt.txt" For Input As #1 'open file
For J = 1 To 31 'fill array
    Input #1, Balance(J)
Next J
picResults.Print "The balance on October"; D; "was "; FormatCurrency(Balance(D)); "."
'print the balance on the day the user requested
Close #1 'close file
End Sub


