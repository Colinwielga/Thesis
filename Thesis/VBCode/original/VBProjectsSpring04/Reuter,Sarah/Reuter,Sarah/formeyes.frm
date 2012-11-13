VERSION 5.00
Begin VB.Form formeyes 
   BackColor       =   &H000000FF&
   Caption         =   "Form2"
   ClientHeight    =   12720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14910
   LinkTopic       =   "Form2"
   ScaleHeight     =   12720
   ScaleWidth      =   14910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnext 
      Caption         =   "Follow me to the next question"
      Enabled         =   0   'False
      Height          =   2055
      Left            =   7200
      Picture         =   "formeyes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton cmdclick 
      Caption         =   "Click to submit answer"
      Height          =   2055
      Left            =   4560
      Picture         =   "formeyes.frx":17F4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   2295
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H008080FF&
      Height          =   4695
      Left            =   840
      ScaleHeight     =   4635
      ScaleWidth      =   8355
      TabIndex        =   5
      Top             =   3360
      Width           =   8415
   End
   Begin VB.OptionButton rdorocks 
      BackColor       =   &H000000FF&
      Caption         =   "Rocks"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   2520
      Width           =   855
   End
   Begin VB.OptionButton rdopaper 
      BackColor       =   &H000000FF&
      Caption         =   "Paper"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   2160
      Width           =   855
   End
   Begin VB.OptionButton rdogolf 
      BackColor       =   &H000000FF&
      Caption         =   "Golf Balls"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.OptionButton rdopingpong 
      BackColor       =   &H000000FF&
      Caption         =   "Ping Pong balls"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Created by Sarah Reuter"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   8280
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "What are Kermit's eyes made out of?"
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "formeyes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program: Sarah's Kermit the Frog Quiz
'Form name and file: formeyes(formeyes.frm)
'Created by Sarah Reuter
'Written 3/14/04
'Purpose: trivia question

Dim Cpingpong As Single
Dim Wgolf As Single
Dim Wpaper As Single
Dim Wrocks As Single

'print correct answer and adjust score
Private Sub cmdclick_Click()
If Cpingpong = 1 Then
    Correct = Correct + 1
    picresults.Print "You're correct, Kermit's eyes are made of ping pong balls."
End If
If Wgolf = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, Kermit's eyes are made of ping pong balls, not golf balls."
End If
If Wpaper = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, Kermit's eyes are made of ping pong balls, not paper."
End If
If Wrocks = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, Kermit's eyes are made of ping pong balls, not rocks."
End If
picresults.Print
picresults.Print "You have gotten"; Wrong; "questions wrong so far."
picresults.Print "You have gotten"; Correct; "questions correct so far."
cmdnext.Enabled = True
cmdclick.Enabled = False
End Sub

'switch to next form
Private Sub cmdnext_Click()
formeyes.Hide
formborn.Show
End Sub

'adjust sub score
Private Sub rdogolf_Click()
Cpingpong = 0
Wgolf = 1
Wpaper = 0
Wrocks = 0
End Sub

Private Sub rdopaper_Click()
Cpingpong = 0
Wgolf = 0
Wpaper = 1
Wrocks = 0
End Sub

Private Sub rdopingpong_Click()
Cpingpong = 1
Wgolf = 0
Wpaper = 0
Wrocks = 0
End Sub

Private Sub rdorocks_Click()
Cpingpong = 0
Wgolf = 0
Wpaper = 0
Wrocks = 1
End Sub
