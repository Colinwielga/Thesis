VERSION 5.00
Begin VB.Form SelectOpponent2 
   BackColor       =   &H0000C000&
   Caption         =   "Form2"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5760
   LinkTopic       =   "Form2"
   ScaleHeight     =   4440
   ScaleWidth      =   5760
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Quit2 
      BackColor       =   &H0000C000&
      Caption         =   "Quit"
      Height          =   615
      Left            =   2880
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Continue2 
      BackColor       =   &H0000C000&
      Caption         =   "Click Here to Continue..."
      Height          =   615
      Left            =   480
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   1935
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H0000FF00&
      Caption         =   "John Ashcroft"
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   2280
      Width           =   3495
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00004000&
      Caption         =   "Charles Manson"
      ForeColor       =   &H0080FF80&
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   3495
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Adolf Hitler"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   3495
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0080FF80&
      Caption         =   "Joseph Stalin"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   3495
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00008000&
      Caption         =   "President George W. Bush"
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "SelectOpponent2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Continue2_Click()
If Option1.Value = True Then
    rival = opponentname(1) 'otherwise just put 1,2,etc in parenths)
    score = opponentfactor(1)
ElseIf Option2.Value = True Then
    rival = opponentname(2)
    score = opponentfactor(2)
ElseIf Option3.Value = True Then
    rival = opponentname(3)
    score = opponentfactor(3)
ElseIf Option4.Value = True Then
    rival = opponentname(4)
    score = opponentfactor(4)
ElseIf Option5.Value = True Then
    rival = opponentname(5)
    score = opponentfactor(5)
End If
SelectOpponent2.Visible = False
qualities23.Visible = True
End Sub

Private Sub Quit2_Click()
End
End Sub
