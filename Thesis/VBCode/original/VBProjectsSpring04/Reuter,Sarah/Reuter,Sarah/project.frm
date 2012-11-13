VERSION 5.00
Begin VB.Form formlive 
   BackColor       =   &H00FF80FF&
   Caption         =   "Form1"
   ClientHeight    =   12615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   ScaleHeight     =   12615
   ScaleWidth      =   14910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnext 
      Caption         =   "Follow me to the next question!"
      Enabled         =   0   'False
      Height          =   2175
      Left            =   7440
      Picture         =   "project.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Width           =   2535
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FFC0FF&
      Height          =   4095
      Left            =   720
      ScaleHeight     =   4035
      ScaleWidth      =   9075
      TabIndex        =   5
      Top             =   4440
      Width           =   9135
   End
   Begin VB.CommandButton cmdclick 
      Caption         =   "Click to submit answer."
      Height          =   2175
      Left            =   3840
      Picture         =   "project.frx":17F4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   2535
   End
   Begin VB.OptionButton rdohollywood 
      BackColor       =   &H00FF80FF&
      Caption         =   "Hollywood"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   3600
      Width           =   1575
   End
   Begin VB.OptionButton rdolondon 
      BackColor       =   &H00FF80FF&
      Caption         =   "London"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.OptionButton rdoantartica 
      BackColor       =   &H00FF80FF&
      Caption         =   "Antartica"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.OptionButton rdoaustrailia 
      BackColor       =   &H00FF80FF&
      Caption         =   "Austrailia"
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF80FF&
      Caption         =   "Created by Sarah Reuter"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   8760
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF80FF&
      Caption         =   "Where does Kermit the Frog live?"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1320
      Width           =   2655
   End
End
Attribute VB_Name = "formlive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program: Sarah's Kermit the Frog Quiz
'Form name and file: formlive(project.frm)
'Created by Sarah Reuter
'Written 3/14/04
'Purpose: trivia question

Dim Waustrailia As Single
Dim Wantartica As Single
Dim Wlondon As Single
Dim Chollywood As Single

'print the answer and adjust the score
Private Sub cmdclick_Click()
If Wantartica = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, but Kermit the Frog does not live in Antartica, he lives in Hollywood."
End If
If Waustrailia = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, but Kermit the Frog does not live in Austrailia, he lives in Hollywood."
End If
If Wlondon = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, but Kermit the Frog does not live in London, he lives in Hollywood."
End If
If Chollywood = 1 Then
    Correct = Correct + 1
    picresults.Print "You're correct!  Kermit the Frog lives in Hollywood!  Good Job!"
End If
picresults.Print
picresults.Print "You have gotten"; Wrong; "questions wrong so far."
picresults.Print "You have gotten"; Correct; "questions correct so far."
cmdnext.Enabled = True
cmdclick.Enabled = False
End Sub

'switch to the next form
Private Sub cmdnext_Click()
formlive.Hide
formsong.Show
End Sub

'adjust the subscores so that the computer knows which answer was chosen
Private Sub rdoantartica_Click()
Wantartica = 1
Waustrailia = 0
Wlondon = 0
Chollywood = 0
End Sub

Private Sub rdoaustrailia_Click()
Waustrailia = 1
Wantartica = 0
Wlondon = 0
Chollywood = 0
End Sub

Private Sub rdohollywood_Click()
Chollywood = 1
Waustrailia = 0
Wantartica = 0
Wlondon = 0
End Sub

Private Sub rdolondon_Click()
Wlondon = 1
Waustrailia = 0
Wantartica = 0
Chollywood = 0
End Sub
