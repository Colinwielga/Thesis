VERSION 5.00
Begin VB.Form frmGrandTotal 
   BackColor       =   &H00FF0000&
   Caption         =   "Grand Total"
   ClientHeight    =   5895
   ClientLeft      =   2100
   ClientTop       =   2445
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdTotalScore 
      BackColor       =   &H0000FFFF&
      Caption         =   "And Your Grand Total Is..."
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1935
   End
   Begin VB.PictureBox picOutput 
      Height          =   2655
      Left            =   2880
      ScaleHeight     =   2595
      ScaleWidth      =   6795
      TabIndex        =   1
      Top             =   2520
      Width           =   6855
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   360
      Picture         =   "frmJ9.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   120
      Width           =   9615
   End
End
Attribute VB_Name = "frmGrandTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Jeopardy.(Jeopardy.vbp)
'Form name: GrandTotal; Form caption: Jeopardy
'Author: Skrbec and Jahnke
'Date written: October 29, 2006
'Form Objective: The objective for this form is to print out the total score that the user
'                ends up with at the end of the game.

Private Sub cmdQuit_Click()
End
End Sub
Private Sub cmdTotalScore_Click()
    Select Case Sum
        Case Is < 0
            picOutput.Print "Your Total Score is "; Sum; ". You Are Not Good At Jeopardy! Quit and try again!"
        Case 0 To 1000
            picOutput.Print "Your Total Score is "; Sum; ". You Need To Try Harder! Quit and try again!"
        Case 1001 To 2000
            picOutput.Print "Your Total Score is "; Sum; ". Better Luck Next Time! Quit and try again!"
        Case 2001 To 3000
            picOutput.Print "Your Total Score is "; Sum; ". You Are Just Average, Keep Trying! Quit and try again!"
        Case 3001 To 4000
            picOutput.Print "Your Total Score is "; Sum; ". Awesome! You Are A Jeopardy Whiz! Quit and try again!"
        Case 4001 To 5000
            picOutput.Print "Your Total Score is "; Sum; ". Wow You're Incredible! You Should Host This Show!"
    End Select                                  '   This button shows the end result of the
End Sub                                         '   the game and how you did. We chose to use
                                                '   a case select method so that it would
                                                '   run through the different cases to see
                                                '   the ending result.
                                                

