VERSION 5.00
Begin VB.Form frmStart 
   Caption         =   "Style Finder"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10590
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   Picture         =   "frmStart.frx":0000
   ScaleHeight     =   8250
   ScaleWidth      =   10590
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H008080FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7080
      Width           =   2535
   End
   Begin VB.CommandButton cmdCelebs 
      BackColor       =   &H008080FF&
      Caption         =   "Read more about your favorite Celebs!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CommandButton cmdTrends 
      BackColor       =   &H008080FF&
      Caption         =   "Trends for Winter 2005/2006!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuiz 
      BackColor       =   &H008080FF&
      Caption         =   "Take Quiz Now!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label lblNames 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenna Gebeke ~ Katie Ranallo"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   9240
      Width           =   3015
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "What's Your Style?"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   9855
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Project Name: What's Your Style?
' Authors: Jenna Gebeke & Katie Ranallo
' Date Written: March 24, 2006
' Overall purpose of the project: To allow the user to determine her style, and a celebrity who shares their style.  The user may also view fashion trends for the season, and get ideas of how to dress within their style.
' This Form Name: Start
' This Form Objective: To allow the user to navigate between the different features of our program including: Take your style quiz, View your favorite celebs,view style trends for Winter 2005/2006,or to exit the program.

Private Sub cmdCelebs_Click()
'This command button allows the user to view her favorite celebs and read about their style.
    userName = InputBox("Before you begin, please enter your name", "User Name")
    frmCelebs.Show
    frmStart.Hide
End Sub

Private Sub cmdExit_Click()
'This command button allows the user to exit the program.
    End
End Sub

Private Sub cmdQuiz_Click()
'This command button allows the user to take a style quiz to determine their style and their celeb style match.
    userName = InputBox("Before you begin, please enter your name", "User Name")
    frmQuiz.Show
    frmStart.Hide
End Sub

Private Sub cmdTrends_Click()
'This command button allows the user to view the fashion trends for Winter 2005/2006.
    userName = InputBox("Before you begin, please enter your name", "User Name")
    frmTrends.Show
    frmStart.Hide
End Sub

