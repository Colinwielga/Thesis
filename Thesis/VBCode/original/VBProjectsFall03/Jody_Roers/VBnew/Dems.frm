VERSION 5.00
Begin VB.Form Dems 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   8670
   ClientLeft      =   870
   ClientTop       =   1395
   ClientWidth     =   13575
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   13575
   Begin VB.PictureBox picDems 
      Height          =   2175
      Left            =   960
      Picture         =   "Dems.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   1875
      TabIndex        =   6
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1215
      Left            =   10560
      TabIndex        =   4
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "What was the Balance?"
      Height          =   1215
      Left            =   7200
      TabIndex        =   3
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Can the College Democrats Club cover the expense?"
      Height          =   1215
      Left            =   3720
      TabIndex        =   2
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Return to the Main Menu"
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   3960
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      Height          =   2175
      Left            =   3840
      ScaleHeight     =   2115
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   840
      Width           =   7935
   End
   Begin VB.Label lblDems 
      BackColor       =   &H00C0FFC0&
      Caption         =   "College Democrats Club"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Dems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Club Auditing Aid (VBProject.vbp)
'Form Name: Comm (VBNew.frm)
'Author: Jody Roers
'Date Written: 27 October 2003
'Purpose: to aid in my responsibilities as a Club Auditor in CSB Senate.  The program can give me the balance for any date in October selected and determine whether the club has enough money for a specificied withdrawal.

Private Sub cmdLoad_Click()
Dim D As Integer
Dim M As Single
Dim J As Integer
Dim Balance(1 To 31) As Single
Dim Sum As Single
picResults.Cls 'clear screen
D = InputBox("What is the October Date?", "Enter Date") 'ask october date user wants to withdraw money on
Do While D > 31 Or D < 1 'if number isn't between 1 and 31(31 days in October) then give error message
    MsgBox "Sorry, you have entered an invalid date", , "Error"
    D = InputBox("What is the October Date?", "Enter Date") 'ask october date again
Loop 'continue until number entered is between 1 and 31
Open "M:\CS130\Projects\VB 10-21-02\Demstxt.txt" For Input As #1 'open file
For J = 1 To 31 'fill array
    Input #1, Balance(J)
Next J
M = InputBox("Subtract how much money?", "Money") 'ask for amount of money to subtract
If Balance(D) > M Then 'checking if club has enough money to subtract that withdrawal
        Sum = Balance(D) - M 'club has enough money and finding new balance in light of withdrawal
        picResults.Print "The College Democrats Club has enough money to cover this. The new balance is "; FormatCurrency(Sum) 'printing new balance
    Else
        picResults.Print "The College Democrats Club does not have enough money to cover this expense." 'letting user know that comm club doesn't have enough money
End If
Close #1 'close file
End Sub


Private Sub cmdMenu_Click()
Dems.Hide 'return to main menu
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
Open "M:\CS130\Projects\VB 10-21-02\Demstxt.txt" For Input As #1 'open file
For J = 1 To 31 'fill array
    Input #1, Balance(J)
Next J
picResults.Print "The balance on October"; D; "was "; FormatCurrency(Balance(D)); "." 'print the balance on the day the user requested
Close #1 'close file
End Sub

