VERSION 5.00
Begin VB.Form frm7th1 
   BackColor       =   &H00000000&
   Caption         =   "7th Chords"
   ClientHeight    =   7995
   ClientLeft      =   3330
   ClientTop       =   2535
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   Picture         =   "frm7th1.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   10695
   Begin VB.CommandButton cmdCheck 
      BackColor       =   &H0080FFFF&
      Caption         =   "CHECK ANSWER"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtAnswer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H008080FF&
      Caption         =   "Next Question"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   2535
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C000&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Label lblOptions 
      BackColor       =   &H80000012&
      Caption         =   "1) Mm7         2) m7             3) M7             4) half dim7 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1935
      Left            =   3840
      TabIndex        =   3
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label lblQuestion 
      BackColor       =   &H80000007&
      Caption         =   "Please Type the # of the Correct Answer:"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   3240
      Width           =   7335
   End
End
Attribute VB_Name = "frm7th1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form corresponds to one of the Seventh Chord questions
'This form corresponds to one of the Seventh Chord questions

Private Sub cmdBack_Click() 'brings the player to the Choose Menu

    frm7th1.Hide
    frmChoose.Show
End Sub

Private Sub cmdCheck_Click() 'asks the player to type # of correct response, then checks.
                             'prints "Correct" in a message box if correct, or "Incorrect" if wrong
                             
Dim response As Double

response = txtAnswer.Text

If response = 2 Then
    MsgBox ("CORRECT!  Click the red button to move on to the next question.")
    ctr = ctr + 1   'adds one to the Public Counter
Else: MsgBox ("Incorrect answer.  The right answer was 2) m7.  Click the red button to move on to the next question.")
End If

    total = total + 1   'adds one to the Public Total
    
    
End Sub

Private Sub cmdNext_Click() 'takes the player to the next question

    frm7th1.Hide
    frm7th2.Show
End Sub
