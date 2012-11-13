VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   Caption         =   "Form1"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   1215
      Left            =   480
      ScaleHeight     =   1155
      ScaleWidth      =   7755
      TabIndex        =   3
      Top             =   4680
      Width           =   7815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdName 
      BackColor       =   &H008080FF&
      Caption         =   "Let's Get Started!!"
      Height          =   975
      Left            =   360
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This program will help high school seniors think about the field
'of study they would like to major in during college.  It will do
'so by asking a series of questions and then match those answers
'with a major that might interest the user.
Dim User As String



Private Sub cmdName_Click()
User = InputBox("Please enter your name:", "Hello!")

picResults.Print "Welcome"; User; "Let's get started and find a major that is right for you!"



End Sub
