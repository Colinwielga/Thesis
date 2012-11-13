VERSION 5.00
Begin VB.Form Fusion 
   BackColor       =   &H80000007&
   Caption         =   "Fusion"
   ClientHeight    =   8370
   ClientLeft      =   465
   ClientTop       =   855
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   13350
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFusion 
      Height          =   2295
      Left            =   240
      Picture         =   "Fusion.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   3315
      TabIndex        =   5
      Top             =   960
      Width           =   3375
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Return to the Main Menu"
      Height          =   1095
      Left            =   480
      TabIndex        =   4
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Can the Cultural Fusion Club cover the expense?"
      Height          =   1095
      Left            =   3120
      TabIndex        =   3
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "What was the balance?"
      Height          =   1095
      Left            =   6240
      TabIndex        =   2
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   9120
      TabIndex        =   1
      Top             =   4200
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      Height          =   2055
      Left            =   4080
      ScaleHeight     =   1995
      ScaleWidth      =   6555
      TabIndex        =   0
      Top             =   1320
      Width           =   6615
   End
   Begin VB.Label lblName 
      Caption         =   "Created by Jody Roers"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblFusion 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cultural Fusion Club"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "Fusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Club Auditing Aid (VBProject.vbp)
'Form Name: Fusion (Fusion.frm)
'Author: Jody Roers
'Date Written: 27 October 2003
'Purpose: to aid in my responsibilities as a Club Auditor in CSB Senate.
'   The program can give me the balance for any date in October selected and
'   determine whether the club has enough money for a specificied withdrawal.


Private Sub cmdLoad_Click()
Dim D As Integer
Dim M As Single
Dim Balance(1 To 31) As Single
Dim J As Integer
Dim Sum As Single
picResults.Cls
D = InputBox("What is the October Date? Only enter the number of the day specified.", "Enter Date") 'ask for october date user wants to withdraw money on
Do While D > 31 Or D < 1 'if number isn't between 1 and 31(31 days in October) then give error message
    MsgBox "Sorry, you have entered an invalid date", , "Error"
    D = InputBox("What is the October Date?", "Enter Date") 'ask october date again
Loop 'continue until number entered is between 1 and 31
Open Menu.Path & "VB 10-21-02\Fusiontxt.txt" For Input As #1 'open file
For J = 1 To 31 'fill array
    Input #1, Balance(J)
Next J
M = InputBox("Subtract how much money?", "Money") 'ask for amount of money to subtract
If Balance(D) > M Then 'checking if club has enough money to subtract that withdrawal
        Sum = Balance(D) - M 'club has enough money and finding new balance in light of withdrawal
        picResults.Print "The Cultural Fusion Club has enough money to cover this. The new balance is "; FormatCurrency(Sum)
        'printing new balance
    Else
        picResults.Print "The Cultural Fusion Club does not have enough money to cover this expense."
        'letting user know that comm club doesn't have enough money
End If
Close #1 'close file
End Sub


Private Sub cmdMenu_Click()
Fusion.Hide 'return to main menu
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
Open Menu.Path & "VB 10-21-02\Fusiontxt.txt" For Input As #1 'open file
For J = 1 To 31 'fill array
    Input #1, Balance(J)
Next J
picResults.Print "The balance on October"; D; "was "; FormatCurrency(Balance(D)); "."
'print the balance on the day the user requested
Close #1 'close file
End Sub


