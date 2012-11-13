VERSION 5.00
Begin VB.Form frmtrivia 
   Caption         =   "Trivia"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   Picture         =   "frmtrivia.frx":0000
   ScaleHeight     =   5190
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picresults 
      BackColor       =   &H0080C0FF&
      Height          =   975
      Left            =   1920
      ScaleHeight     =   915
      ScaleWidth      =   3675
      TabIndex        =   2
      Top             =   3240
      Width           =   3735
   End
   Begin VB.CommandButton cmdmain 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Main Menu"
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdquiz 
      BackColor       =   &H0080FFFF&
      Caption         =   "Click Here To Begin!"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   1335
   End
End
Attribute VB_Name = "frmtrivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'A Day For Fun
    'Trivia
    'Stephanie Fiecke
    '10-31-06
    'This form is for trivia fun and allows the uswer to answer a few questions
    'the program keeps track of their score and gives them a total at the end of the game.
    
Option Explicit
Private Sub cmdmain_Click()
    'returns the user to the main menu
frmtrivia.Hide
frmmain.Show
End Sub
Private Sub cmdquiz_Click()
Dim answer1 As String, answer2 As String, answer3 As String, answer4 As String, answer5 As String
Dim score As Integer

    'User is able to put answers into a textbox and be notified if their answer is correct
    'or incorrect by message boxes.

score = 0

 answer1 = InputBox("What is the state bird of Minnesota?")
    If answer1 = "loon" Or answer1 = "Loon" Then
        MsgBox "That is the correct answer!", , "Congratulations!"
        score = score + 1
    Else
        MsgBox "That is not the correct answer!", , "Wrong!"
    End If
 answer2 = InputBox("True or False. A duck's quack echos?")
    If answer2 = "true" Or answer2 = "True" Then
        MsgBox "That is the correct answer!", , "Congratulations!"
        score = score + 1
    Else
        MsgBox "That is not the correct answer!", , "Wrong!"
    End If
answer3 = InputBox("Choose the correct answer. A zebra is white with black stripes or black with white stripes?")
    If answer3 = "white with black stripes" Or answer3 = "White with black stripes" Then
        MsgBox "That is the correct answer!", , "Congratulations!"
        score = score + 1
    Else
        MsgBox "That is not the correct answer!", , "Wrong!"
    End If
answer4 = InputBox("True or False. An ostrich's brain is bigger than its eye.")
    If answer4 = "false" Or answer4 = "False" Then
        MsgBox "That is the correct answer!", , "Congratulations!"
        score = score + 1
    Else
        MsgBox "That is not the correct answer!", , "Wrong!"
    End If
answer5 = InputBox("True or False. By feeding hens certain dyes, they can be made to lay eggs with varicolored yolks.")
    If answer5 = "true" Or answer5 = "True" Then
        MsgBox "That is the correct answer!", , "Congratulations!"
        score = score + 1
    Else
        MsgBox "That is not the correct answer!", , "Wrong!"
    End If
    
    'prints the final score in a picture box
picresults.Print "Your Final Score is"; score; "Good Game!"

End Sub
