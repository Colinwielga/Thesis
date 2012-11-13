VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   Picture         =   "Gilligan's Island.frx":0000
   ScaleHeight     =   11040
   ScaleWidth      =   13545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgame 
      Caption         =   "Can you recognize a face?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   10560
      TabIndex        =   5
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton cmdsources 
      Caption         =   "Sources"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8280
      TabIndex        =   4
      Top             =   7200
      Width           =   2055
   End
   Begin VB.CommandButton cmdorder 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   10560
      TabIndex        =   3
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton cmdquiz 
      Caption         =   "Which character are you?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6000
      TabIndex        =   2
      Top             =   7200
      Width           =   2055
   End
   Begin VB.CommandButton cmdisland 
      Caption         =   "Get to know the island!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8280
      TabIndex        =   1
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton cmdcharacters 
      Caption         =   "Get to know the characters!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6000
      TabIndex        =   0
      Top             =   5400
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project name: Gilligan's Island
'Form name:  Gilligan's Island
'Author:  Emily Olson
'Date written:  March 19, 2008
'Overall Objective: inform user about the classic hit show "Gilligan's Island"

Private Sub cmdgame_Click()
    Form6.Show
    Form1.Hide
End Sub

Private Sub cmdisland_Click()
'load form 3
    Form1.Hide
    Form3.Show
End Sub

Private Sub cmdorder_Click()
'quit
    End
End Sub

Private Sub cmdquiz_Click()
'load form 4
    Form1.Hide
    Form4.Show
End Sub

Private Sub cmdsources_Click()
'load form 5
    Form1.Hide
    Form5.Show
End Sub

Private Sub Form_Load()
    user = InputBox("Hi! What is your name?", "Welcome!")
    MsgBox "Welcome " & user & "!  You are now stranded on Gilligan's Island!!", vbOKOnly, "Welcome!"
'start with first form
    Form1.Show
End Sub


Private Sub cmdcharacters_Click()
'load form 2
    Form2.Show
    Form1.Hide
End Sub

