VERSION 5.00
Begin VB.Form frmKeys4 
   BackColor       =   &H80000007&
   Caption         =   "Key Signatures"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   Picture         =   "frmKeys4.frx":0000
   ScaleHeight     =   7515
   ScaleWidth      =   10380
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
      Left            =   8880
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
      Left            =   7800
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
      Left            =   7680
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
      Left            =   1200
      TabIndex        =   2
      Top             =   2880
      Width           =   7335
   End
   Begin VB.Label lblOptions 
      BackColor       =   &H80000012&
      Caption         =   "1) F# 2) A#  3) E#  4) C#"
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
      Height          =   2175
      Left            =   4440
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "frmKeys4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form corresponds to one of the Key Signature questions

Private Sub cmdBack_Click() 'brings the player to the Choose Menu
    frmKeys4.Hide
    frmChoose.Show
End Sub

Private Sub cmdCheck_Click() 'asks the player to type # of correct response, then checks.
                             'prints "Correct" in a message box if correct, or "Incorrect" if wrong
                             
Dim response As Double

response = txtAnswer.Text

If response = 1 Then
    MsgBox ("CORRECT!  Click the red button to move on to the next question.")
    ctr = ctr + 1   'adds one to the Public Counter
Else: MsgBox ("Incorrect answer.  The right answer was 1) F#.  Click the red button to move on to the next question.")
End If

    total = total + 1   'adds one to the Public Total
    

End Sub
