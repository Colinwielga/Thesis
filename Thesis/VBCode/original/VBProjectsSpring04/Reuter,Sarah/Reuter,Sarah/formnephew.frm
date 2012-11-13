VERSION 5.00
Begin VB.Form formnephew 
   BackColor       =   &H0000C0C0&
   Caption         =   "Form2"
   ClientHeight    =   12645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14865
   LinkTopic       =   "Form2"
   ScaleHeight     =   12645
   ScaleWidth      =   14865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdscore 
      Caption         =   "Follow me to get your score!"
      Enabled         =   0   'False
      Height          =   2175
      Left            =   6480
      Picture         =   "formnephew.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton cmdclick 
      Caption         =   "Click to submit answer"
      Height          =   2175
      Left            =   3840
      Picture         =   "formnephew.frx":17F4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   2295
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H0000FFFF&
      Height          =   3735
      Left            =   600
      ScaleHeight     =   3675
      ScaleWidth      =   8475
      TabIndex        =   5
      Top             =   4320
      Width           =   8535
   End
   Begin VB.OptionButton rdofrancis 
      BackColor       =   &H0000C0C0&
      Caption         =   "Francis"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.OptionButton rdorobin 
      BackColor       =   &H0000C0C0&
      Caption         =   "Robin"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   2640
      Width           =   855
   End
   Begin VB.OptionButton rdojesse 
      BackColor       =   &H0000C0C0&
      Caption         =   "Jesse"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2160
      Width           =   855
   End
   Begin VB.OptionButton rdoandrew 
      BackColor       =   &H0000C0C0&
      Caption         =   "Andrew"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "Created by Sarah Reuter"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   8280
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Who is Kermit's nephew?"
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "formnephew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program: Sarah's Kermit the Frog Quiz
'Form name and file: formnephew(formnephew.frm)
'Created by Sarah Reuter
'Written 3/14/04
'Purpose: trivia question

Dim Wandrew As Single
Dim Wjesse As Single
Dim Crobin As Single
Dim Wfrancis As Single

'print correct answer and adjust score
Private Sub cmdclick_Click()
If Wandrew = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, Kermit's nephew's name is Robin, not Andrew."
End If
If Wjesse = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, Kermit's nephew's name is Robin, not Jesse."
End If
If Crobin = 1 Then
    Correct = Correct + 1
    picresults.Print "You're correct.  Kermit's nephew's name is Robin."
End If
If Wfrancis = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, Kermit's nephew's name is Francis."
End If
picresults.Print
picresults.Print "You have gotten"; Wrong; "questions wrong so far."
picresults.Print "You have gotten"; Correct; "questions correct so far."
cmdscore.Enabled = True
cmdclick.Enabled = False
End Sub

'switch to next form
Private Sub cmdscore_Click()
formnephew.Hide
formscore.Show
End Sub

'adjust sub score
Private Sub rdoandrew_Click()
Wandrew = 1
Wjesse = 0
Crobin = 0
Wfrancis = 0
End Sub

Private Sub rdofrancis_Click()
Wandrew = 0
Wjesse = 0
Crobin = 0
Wfrancis = 1
End Sub

Private Sub rdojesse_Click()
Wandrew = 0
Wjesse = 1
Crobin = 0
Wfrancis = 0
End Sub

Private Sub rdorobin_Click()
Wandrew = 0
Wjesse = 0
Crobin = 1
Wfrancis = 0
End Sub
