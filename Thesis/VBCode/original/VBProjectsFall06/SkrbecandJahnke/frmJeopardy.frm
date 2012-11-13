VERSION 5.00
Begin VB.Form frmJeopardy 
   Caption         =   "Jeopardy"
   ClientHeight    =   9585
   ClientLeft      =   1470
   ClientTop       =   915
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   ScaleHeight     =   9585
   ScaleWidth      =   11985
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picJeopardy 
      Height          =   9615
      Left            =   0
      Picture         =   "frmJeopardy.frx":0000
      ScaleHeight     =   9555
      ScaleWidth      =   11955
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   9120
         TabIndex        =   2
         Top             =   8640
         Width           =   2175
      End
      Begin VB.CommandButton cmdEnter 
         Caption         =   "Enter Game"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   2400
         TabIndex        =   1
         Top             =   6960
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmJeopardy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EnterName As String
'Jeopardy.(Jeopardy.vbp)
'Form name: Jeopardy; Form caption: Jeopardy
'Author: Skrbec and Jahnke
'Date written: October 29, 2006
'Form Objective: This is the main form that first shows up when the game is started.
'                After it asks for the users name, it then asks them if they are ready
'                to play the game. This is a simple and basic starting form which then
'                leads them into the actual game (a new form) where they can start to choose different
'                categories and whatnot.

'This button allows the user to enter their name in which it then asks if they are
'ready to play the game.

Private Sub cmdEnter_Click()
EnterName = InputBox("Enter Your Name to Log In", "Enter Name")
MsgBox "Welcome to Jeopardy! Are you ready to play" & " " & EnterName & "?", , "Welcome"
        frmJeopardy.Hide
        frmTopics.Show
        
End Sub
'This is the standard quit button.
Private Sub cmdQuit_Click()
End
End Sub

