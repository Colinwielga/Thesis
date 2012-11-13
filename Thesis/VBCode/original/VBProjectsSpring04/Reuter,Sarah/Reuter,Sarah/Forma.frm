VERSION 5.00
Begin VB.Form formcreate 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   12705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14940
   LinkTopic       =   "Form1"
   ScaleHeight     =   12705
   ScaleWidth      =   14940
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picresults 
      BackColor       =   &H0080C0FF&
      Height          =   4095
      Left            =   600
      ScaleHeight     =   4035
      ScaleWidth      =   8955
      TabIndex        =   7
      Top             =   4320
      Width           =   9015
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Follow me to the next question"
      Enabled         =   0   'False
      Height          =   2175
      Left            =   7200
      Picture         =   "Forma.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton cmdclick 
      Caption         =   "Click to submit answer"
      Height          =   2175
      Left            =   4200
      Picture         =   "Forma.frx":17F4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   2295
   End
   Begin VB.OptionButton rdojim 
      BackColor       =   &H000080FF&
      Caption         =   "Jim Henson"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   3480
      Width           =   1335
   End
   Begin VB.OptionButton rdocharles 
      BackColor       =   &H000080FF&
      Caption         =   "Charles O'Hera"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.OptionButton rdobob 
      BackColor       =   &H000080FF&
      Caption         =   "Bob Barker"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.OptionButton rdomichael 
      BackColor       =   &H000080FF&
      Caption         =   "Michael Jackson"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Caption         =   "Created by Sarah Reuter"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   8640
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "Who created Kermit the Frog?"
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   2295
   End
End
Attribute VB_Name = "formcreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program: Sarah's Kermit the Frog Quiz
'Form name and file: formcreate(forma.frm)
'Created by Sarah Reuter
'Written 3/14/04
'Purpose: trivia question

Dim Wmichael As Single
Dim Wbob As Single
Dim Wcharles As Single
Dim Cjim As Single

'print correct answer and adjust score
Private Sub cmdclick_Click()
If Wmichael = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, Jim Henson created Kermit the frog not Michael Jackson."
End If
If Wbob = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, Jim Henson created Kermit the frog, not Bob Barker."
End If
If Wcharles = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, Jim Henson created Kermit the frog, not Charles O'Hera."
End If
If Cjim = 1 Then
    Correct = Correct + 1
    picresults.Print "You're correct, Jim Henson created Kermit the frog."
End If
picresults.Print
picresults.Print "You have gotten"; Wrong; "questions wrong so far."
picresults.Print "You have gotten"; Correct; "questions correct so far."
cmdnext.Enabled = True
cmdclick.Enabled = False
End Sub

'switch to next form
Private Sub cmdnext_Click()
formcreate.Hide
formeyes.Show
End Sub

'adjust sub score
Private Sub rdobob_Click()
Wmichael = 0
Wbob = 1
Wcharles = 0
Cjim = 0
End Sub

Private Sub rdocharles_Click()
Wmichael = 0
Wbob = 0
Wcharles = 1
Cjim = 0
End Sub

Private Sub rdojim_Click()
Wmichael = 0
Wbob = 0
Wcharles = 0
Cjim = 1
End Sub

Private Sub rdomichael_Click()
Wmichael = 1
Wbob = 0
Wcharles = 0
Cjim = 0
End Sub
