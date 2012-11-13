VERSION 5.00
Begin VB.Form formwhere 
   BackColor       =   &H00C000C0&
   Caption         =   "Form2"
   ClientHeight    =   12645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14865
   LinkTopic       =   "Form2"
   ScaleHeight     =   12645
   ScaleWidth      =   14865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnext 
      Caption         =   "Follow me to the next question"
      Enabled         =   0   'False
      Height          =   2055
      Left            =   6840
      Picture         =   "formwhere.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdclick 
      Caption         =   "Click to submit answer"
      Height          =   2055
      Left            =   3960
      Picture         =   "formwhere.frx":17F4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   2295
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FF00FF&
      Height          =   3495
      Left            =   600
      ScaleHeight     =   3435
      ScaleWidth      =   8835
      TabIndex        =   5
      Top             =   4080
      Width           =   8895
   End
   Begin VB.OptionButton rdotennessee 
      BackColor       =   &H00C000C0&
      Caption         =   "in the hills of Tennessee"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   3240
      Width           =   2175
   End
   Begin VB.OptionButton rdominnesota 
      BackColor       =   &H00C000C0&
      Caption         =   "in the lakes of Minnesota"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   2760
      Width           =   2175
   End
   Begin VB.OptionButton rdonewyork 
      BackColor       =   &H00C000C0&
      Caption         =   "in the streets of New York"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   2280
      Width           =   2175
   End
   Begin VB.OptionButton rdogeorgia 
      BackColor       =   &H00C000C0&
      Caption         =   "in the swamps of Georgia"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C000C0&
      Caption         =   "Created by Sarah Reuter"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C000C0&
      Caption         =   "Where was Kermit born?"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "formwhere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program: Sarah's Kermit the Frog Quiz
'Form name and file: formwhere(formwhere.frm)
'Created by Sarah Reuter
'Written 3/14/04
'Purpose: trivia question

Dim Cgeorgia As Single
Dim Wnewyork As Single
Dim Wminnesota As Single
Dim Wtennessee As Single


'display correct answer and adjust score
Private Sub cmdclick_Click()
If Cgeorgia = 1 Then
    Correct = Correct + 1
    picresults.Print "You're correct, Kermit was born in the swamps of Georgia."
End If
If Wnewyork = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, Kermit was born in the swamps of Georgia, not in the streets of New York."
End If
If Wminnesota = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, Kermit was born in the swamps of Georgia, not in the lakes of Minnesota."
End If
If Wtennessee = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, Kermit was born in the swamps of Georgia, not in the hills of Tennessee."
End If
picresults.Print
picresults.Print
picresults.Print "You have gotten"; Wrong; "questions wrong so far."
picresults.Print "You have gotten"; Correct; "questions correct so far."
cmdnext.Enabled = True
cmdclick.Enabled = False
End Sub

'switch to next form
Private Sub cmdnext_Click()
formwhere.Hide
formnephew.Show
End Sub

'adjust sub score
Private Sub rdogeorgia_Click()
Cgeorgia = 1
Wnewyork = 0
Wminnesota = 0
Wtennessee = 0
End Sub

Private Sub rdominnesota_Click()
Cgeorgia = 0
Wnewyork = 0
Wminnesota = 1
Wtennessee = 0
End Sub

Private Sub rdonewyork_Click()
Cgeorgia = 0
Wnewyork = 1
Wminnesota = 0
Wtennessee = 0
End Sub

Private Sub rdotennessee_Click()
Cgeorgia = 0
Wnewyork = 0
Wminnesota = 0
Wtennessee = 1
End Sub
