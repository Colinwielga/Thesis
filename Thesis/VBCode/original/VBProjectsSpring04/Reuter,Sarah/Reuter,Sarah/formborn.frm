VERSION 5.00
Begin VB.Form formborn 
   BackColor       =   &H00C0C000&
   Caption         =   "Form2"
   ClientHeight    =   12750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15945
   LinkTopic       =   "Form2"
   ScaleHeight     =   12750
   ScaleWidth      =   15945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnext 
      Caption         =   "Follow me to the next question"
      Enabled         =   0   'False
      Height          =   2175
      Left            =   6600
      Picture         =   "formborn.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdclick 
      Caption         =   "Click to submit answer"
      Height          =   2175
      Left            =   3840
      Picture         =   "formborn.frx":17F4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   2295
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FFFFC0&
      Height          =   3975
      Left            =   600
      ScaleHeight     =   3915
      ScaleWidth      =   8475
      TabIndex        =   5
      Top             =   4080
      Width           =   8535
   End
   Begin VB.OptionButton rdo1964 
      BackColor       =   &H00C0C000&
      Caption         =   "1964"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   3120
      Width           =   735
   End
   Begin VB.OptionButton rdo1955 
      BackColor       =   &H00C0C000&
      Caption         =   "1955"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   2640
      Width           =   735
   End
   Begin VB.OptionButton rdo1938 
      BackColor       =   &H00C0C000&
      Caption         =   "1938"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   2160
      Width           =   735
   End
   Begin VB.OptionButton rdo1895 
      BackColor       =   &H00C0C000&
      Caption         =   "1895"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Created by Sarah Reuter"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   8280
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "When was Kermit born?"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "formborn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program: Sarah's Kermit the Frog Quiz
'Form name and file: formborn(formborn.frm)
'Created by Sarah Reuter
'Written 3/14/04
'Purpose: trivia question

Dim W1895 As Single
Dim W1938 As Single
Dim C1955 As Single
Dim W1964 As Single

'print the correct answer and adjust score
Private Sub cmdclick_Click()
If W1895 = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, Kermit was born in 1955, not 1895."
End If
If W1938 = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, Kermit was born in 1955, not 1938."
End If
If C1955 = 1 Then
    Correct = Correct + 1
    picresults.Print "You're correct, Kermit was born in 1955."
End If
If W1964 = 1 Then
    Wrong = Wrong + 1
    picresults.Print "I'm sorry, Kermit was born in 1955, not 1964."
End If
picresults.Print
picresults.Print "You have gotten"; Wrong; "questions wrong so far."
picresults.Print "You have gotten"; Correct; "questions correct so far."
cmdnext.Enabled = True
cmdclick.Enabled = False
End Sub

'switch to next form
Private Sub cmdnext_Click()
formborn.Hide
formwhere.Show
End Sub

'adjust subscore
Private Sub rdo1895_Click()
W1895 = 1
W1938 = 0
C1955 = 0
W1964 = 0
End Sub

Private Sub rdo1938_Click()
W1895 = 0
W1938 = 1
C1955 = 0
W1964 = 0
End Sub

Private Sub rdo1955_Click()
W1895 = 0
W1938 = 0
C1955 = 1
W1964 = 0
End Sub

Private Sub rdo1964_Click()
W1985 = 0
W1938 = 0
C1955 = 0
W1964 = 1
End Sub
