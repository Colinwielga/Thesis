VERSION 5.00
Begin VB.Form frmTriads4 
   BackColor       =   &H80000012&
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   Picture         =   "frmTriads4.frx":0000
   ScaleHeight     =   7710
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
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
      Left            =   7440
      TabIndex        =   3
      Top             =   2760
      Width           =   855
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   2535
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
      Left            =   840
      TabIndex        =   2
      Top             =   2880
      Width           =   7335
   End
   Begin VB.Label lblOptions 
      BackColor       =   &H80000012&
      Caption         =   "1)  A minor                   2)  A major                   3)  A augmented          4)  A diminished"
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
      Height          =   2055
      Left            =   3600
      TabIndex        =   1
      Top             =   3480
      Width           =   3495
   End
End
Attribute VB_Name = "frmTriads4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form corresponds to one of the Triad questions

Private Sub cmdBack_Click() 'brings the player to the Choose Menu
    frmTriads4.Hide
    frmChoose.Show
End Sub


Private Sub cmdCheck_Click() 'asks the player to type # of correct response, then checks.
                             'prints "Correct" in a message box if correct, or "Incorrect" if wrong
                             
Dim response As Double

response = txtAnswer.Text

If response = 3 Then
    MsgBox ("CORRECT!  Click the red button to move on to the next question.")
    ctr = ctr + 1   'adds one to the Public Counter
Else: MsgBox ("Incorrect answer.  The right answer was 3) A Augmented.  Click the red button to move on to the next question.")
End If

    total = total + 1   'adds one to the Public Total
    

End Sub
