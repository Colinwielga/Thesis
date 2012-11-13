VERSION 5.00
Begin VB.Form formphrase 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form1"
   ClientHeight    =   12690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   12690
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picresults 
      BackColor       =   &H0080FFFF&
      Height          =   4215
      Left            =   720
      ScaleHeight     =   4155
      ScaleWidth      =   8475
      TabIndex        =   7
      Top             =   3960
      Width           =   8535
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Follow me to the next question"
      Enabled         =   0   'False
      Height          =   2055
      Left            =   7560
      Picture         =   "formphrase.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdclick 
      Caption         =   "Click to submit answer"
      Height          =   2055
      Left            =   4920
      Picture         =   "formphrase.frx":17F4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   2175
   End
   Begin VB.OptionButton rdotaste 
      BackColor       =   &H0000FFFF&
      Caption         =   "It taste like burning."
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.OptionButton rdohiho 
      BackColor       =   &H0000FFFF&
      Caption         =   "Hi Ho!"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin VB.OptionButton rdolook 
      BackColor       =   &H0000FFFF&
      Caption         =   "Look both ways before crossing the street"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1920
      Width           =   3375
   End
   Begin VB.OptionButton rdosuper 
      BackColor       =   &H0000FFFF&
      Caption         =   "Supercalafrogalisticexpiealadocious"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Created by Sarah Reuter"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   8400
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "What is Kermit's favorite phrase?"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
End
Attribute VB_Name = "formphrase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program: Sarah's Kermit the Frog Quiz
'Form name and file: formphrase(formphrase.frm)
'Created by Sarah Reuter
'Written 3/14/04
'Purpose: trivia question

Dim Wsuper As Single
Dim Chiho As Single
Dim Wlook As Single
Dim Wtaste As Single

'print correct answer and adjust score
Private Sub cmdclick_Click()
If Wsuper = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, Kermit says Hi Ho, not Supercalafrogalisticexpialadocious"
End If
If Chiho = 1 Then
    Correct = Correct + 1
    picresults.Print "You're correct!  Kermit the frog says Hi Ho!"
End If
If Wlook = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, Kermit says Hi Ho, Not Look both ways before crossing the street"
End If
If Wtaste = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, Kermit says Hi Ho, not It tastes like burning."
End If
picresults.Print
picresults.Print "You have gotten"; Wrong; "questions wrong so far."
picresults.Print "You have gotten"; Correct; "questions correct so far."
cmdnext.Enabled = True
cmdclick.Enabled = False
End Sub

'switch to next form
Private Sub cmdnext_Click()
formphrase.Hide
forminstrument.Show
End Sub

'adjust sub score
Private Sub rdohiho_Click()
Wsuper = 0
Chiho = 1
Wlook = 0
Wtaste = 0
End Sub

Private Sub rdolook_Click()
Wsuper = 0
Chiho = 0
Wlook = 1
Wtaste = 0
End Sub

Private Sub rdosuper_Click()
Wsuper = 1
Chiho = 0
Wlook = 0
Wtaste = 0
End Sub

Private Sub rdotaste_Click()
Wsuper = 0
Chiho = 0
Wlook = 0
Wtaste = 1
End Sub
