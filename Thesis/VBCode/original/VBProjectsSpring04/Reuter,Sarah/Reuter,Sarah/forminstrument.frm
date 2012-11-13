VERSION 5.00
Begin VB.Form forminstrument 
   BackColor       =   &H0000FF00&
   Caption         =   "Form1"
   ClientHeight    =   12645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14895
   LinkTopic       =   "Form1"
   ScaleHeight     =   12645
   ScaleWidth      =   14895
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picresults 
      BackColor       =   &H0080FF80&
      Height          =   4575
      Left            =   720
      ScaleHeight     =   4515
      ScaleWidth      =   8595
      TabIndex        =   7
      Top             =   3840
      Width           =   8655
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Follow me to the next question"
      Enabled         =   0   'False
      Height          =   2175
      Left            =   7440
      Picture         =   "forminstrument.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton cmdclick 
      Caption         =   "Click to submit answer"
      Height          =   2175
      Left            =   4200
      Picture         =   "forminstrument.frx":17F4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   2535
   End
   Begin VB.OptionButton rdotrumpet 
      BackColor       =   &H0000FF00&
      Caption         =   "Trumpet"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.OptionButton rdobanjo 
      BackColor       =   &H0000FF00&
      Caption         =   "Banjo"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   2520
      Width           =   855
   End
   Begin VB.OptionButton rdobagpipes 
      BackColor       =   &H0000FF00&
      Caption         =   "Bag pipes"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.OptionButton rdoflute 
      BackColor       =   &H0000FF00&
      Caption         =   "Flute"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "Created by Sarah Reuter"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   8640
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "What instrument does Kermit play?"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   2535
   End
End
Attribute VB_Name = "forminstrument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program: Sarah's Kermit the Frog Quiz
'Form name and file: forminstrument(forminstrument.frm)
'Created by Sarah Reuter
'Written 3/14/04
'Purpose: trivia question

Dim Wflute As Single
Dim Cbanjo As Single
Dim Wtrumpet As Single
Dim Wbagpipes As Single

'print correct answer and adjust score
Private Sub cmdclick_Click()
If Wflute = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, Kermit plays the banjo, not the flute."
End If
If Cbanjo = 1 Then
    Correct = Correct + 1
    picresults.Print "You're correct.  Kermit plays the banjo."
End If
If Wtrumpet = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, Kermit plays the banjo, not the trumpet."
End If
If Wbagpipes = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, Kermit plays the banjo, not the bagpipes."
End If
picresults.Print
picresults.Print "You have gotten"; Wrong; "questions wrong so far."
picresults.Print "You have gotten"; Correct; "questions correct so far."
cmdnext.Enabled = True
cmdclick.Enabled = False
End Sub

'switch to next form
Private Sub cmdnext_Click()
forminstrument.Hide
formcreate.Show
End Sub

'adjust sub score
Private Sub rdobagpipes_Click()
Wflute = 0
Cbanjo = 0
Wtrumpet = 0
Wbagpipes = 1
End Sub

Private Sub rdobanjo_Click()
Wflute = 0
Cbanjo = 1
Wtrumpet = 0
Wbagpipes = 0
End Sub

Private Sub rdoflute_Click()
Wflute = 1
Cbanjo = 0
Wtrumpet = 0
Wbagpipes = 0
End Sub

Private Sub rdotrumpet_Click()
Wflute = 0
Cbanjo = 0
Wtrumpet = 1
Wbagpipes = 0
End Sub
