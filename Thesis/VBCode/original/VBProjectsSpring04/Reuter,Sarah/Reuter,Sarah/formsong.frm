VERSION 5.00
Begin VB.Form formsong 
   BackColor       =   &H00FF8080&
   Caption         =   "Form1"
   ClientHeight    =   12705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   ScaleHeight     =   12705
   ScaleWidth      =   14910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnext 
      Caption         =   "Follow me to the next question."
      Enabled         =   0   'False
      Height          =   2175
      Left            =   6840
      Picture         =   "formsong.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton cmdclick 
      Caption         =   "Click to submit answer."
      Height          =   2175
      Left            =   4200
      Picture         =   "formsong.frx":17F4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   2415
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FFC0C0&
      Height          =   4695
      Left            =   600
      ScaleHeight     =   4635
      ScaleWidth      =   7155
      TabIndex        =   5
      Top             =   3360
      Width           =   7215
   End
   Begin VB.OptionButton rdostar 
      BackColor       =   &H00FF8080&
      Caption         =   "The Star Spangled Banner"
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   2400
      Width           =   2415
   End
   Begin VB.OptionButton rdomary 
      BackColor       =   &H00FF8080&
      Caption         =   "Mary had a Little Lamb"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin VB.OptionButton rdorainbow 
      BackColor       =   &H00FF8080&
      Caption         =   "The Rainbow Connection"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1560
      Width           =   2295
   End
   Begin VB.OptionButton rdobarney 
      BackColor       =   &H00FF8080&
      Caption         =   "The Barney Theme Song"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Created by Sarah Reuter"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   8280
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "What song does Kermit like to sing?"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "formsong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program: Sarah's Kermit the Frog Quiz
'Form name and file: formsong(formsong.frm)
'Created by Sarah Reuter
'Written 3/14/04
'Purpose: trivia question

Dim Wbarney As Single
Dim Wmary As Single
Dim Wstar As Single
Dim Crainbow As Single

'display correct answer and adjust score
Private Sub cmdclick_Click()
If Wbarney = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, but Kermit doesn't sing the Barney theme song, he sings The Rainbow Connection."
End If
If Wmary = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, but Kermit doesn't sing Mary had a Little Lamb, he sings The Rainbow Connection."
End If
If Wstar = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, but Kermit doesn't sing the Star Spangled Banner, he sings The Rainbow Connection."
End If
If Crainbow = 1 Then
    Correct = Correct + 1
    picresults.Print "You're Correct!  Kermit the Frog sings The Rainbow Connection!"
End If
picresults.Print
picresults.Print "You have gotten"; Wrong; "questions wrong so far."
picresults.Print "You have gotten"; Correct; "questions correct so far."
cmdnext.Enabled = True
cmdclick.Enabled = False
End Sub

'switch to next form
Private Sub cmdnext_Click()
formsong.Hide
formphrase.Show
End Sub

'adjust sub score
Private Sub rdobarney_Click()
Wbarney = 1
Wmary = 0
Wstar = 0
Crainbow = 0
End Sub

Private Sub rdomary_Click()
Wbarney = 0
Wmary = 1
Wstar = 0
Crainbow = 0
End Sub

Private Sub rdorainbow_Click()
Wbarney = 0
Wmary = 0
Wstar = 0
Crainbow = 1
End Sub

Private Sub rdostar_Click()
Wstar = 1
Wmary = 0
Wbarney = 0
Crainbow = 0
End Sub
