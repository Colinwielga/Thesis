VERSION 5.00
Begin VB.Form frmFinal 
   BackColor       =   &H8000000D&
   Caption         =   "Final Jeopardy"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12540
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   12540
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picWager 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      ScaleHeight     =   555
      ScaleWidth      =   2115
      TabIndex        =   11
      Top             =   2040
      Width           =   2175
   End
   Begin VB.PictureBox picTotal 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   10
      Top             =   2400
      Width           =   1095
   End
   Begin VB.PictureBox picContestant 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      ScaleHeight     =   435
      ScaleWidth      =   2475
      TabIndex        =   9
      Top             =   1080
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   240
      Picture         =   "frmFinal.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   4755
      TabIndex        =   8
      Top             =   3600
      Width           =   4815
   End
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "Go back to Main Menu without saving"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6480
      TabIndex        =   7
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit game without saving"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9240
      TabIndex        =   6
      Top             =   6600
      Width           =   1695
   End
   Begin VB.PictureBox picFinal 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   5520
      ScaleHeight     =   3195
      ScaleWidth      =   6795
      TabIndex        =   5
      Top             =   3120
      Width           =   6855
   End
   Begin VB.CommandButton cmdClue 
      BackColor       =   &H0080FF80&
      Caption         =   "Get your Final Jeopardy Question Here!"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox txtWager 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Text            =   "Your Wager:"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdWager 
      BackColor       =   &H0080FF80&
      Caption         =   "Click Here to Place Wager"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00FF00FF&
      Caption         =   "    Total:"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF00FF&
      Caption         =   "Contestant:"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: CSB/SJU Jeopardy
'Form Name: frmFinal
'Authors: Emma Jaynes, Lindsay Havlik, Brooke Beyer
'Date Written: 11/02/08
'Objective: the purpose of this form is to ask the user for a wager that is less than or equal to their total and then ask them the Final Jeopardy question.  After the question
'   is asked, the final score is displayed in a picture box as well as a nice message
'Comments: the contestant name and total are carried over from the previous form.  There is a quit button, a button linking us to the main menu, a wager button, and a question button.
'   These buttons do exactly as they are labeled.

Option Explicit
Dim F As String, Wager As Single

Private Sub cmdClue_Click()
'ask user for the answer to the final jeopardy question
F = InputBox("This man is the best CSCI 130 professor in the galaxy. Who is...")

If LCase(F) = "josh trutwin" Then   'these are the two options that are possible with either a right or wrong answer
    MsgBox ("WOW!!! You are so smart! Congratulations!")
    Total = Total + Wager
Else
    MsgBox ("Really? I don't think so, the answer is obviously Josh Trutwin!!")
    Total = Total - Wager
End If

'print results in the picture box
picFinal.Print "Your Grand Total is:"; FormatCurrency(Total); ""
picFinal.Print "We hope you enjoyed playing Jeopardy!"
picFinal.Print "Please play again soon! :)"

End Sub

Private Sub cmdMainMenu_Click()
'opens the main menu form while closing the final jeopardy form and clearing the contestant box and total box
frmFinal.Hide
frmMainMenu.Show
picContestant.Cls
picTotal.Cls

End Sub

Private Sub cmdQuit_Click()
'quits the game
End
End Sub

Private Sub cmdWager_Click()
'asks user for wager using an input box
MsgBox ("Welcome to Final Jeopardy!")
Wager = InputBox("Please enter your wager:")

If Wager > Total Then
    MsgBox ("Wager must not be more than your total, enter another wager:")
    Wager = InputBox("Please enter your wager:")
End If

picWager.Print FormatCurrency(Wager)

'enabled the final jeopardy question button so they can click on it following this step
cmdClue.Enabled = True

End Sub

