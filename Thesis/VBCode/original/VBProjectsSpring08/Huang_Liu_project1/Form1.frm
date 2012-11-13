VERSION 5.00
Begin VB.Form frmgame 
   Caption         =   "Game"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6795
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit "
      Height          =   615
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back to main page"
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton CMD_1 
      Caption         =   "1"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
   Begin VB.Label LBL_2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Range"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.Label LBL_1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmgame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RAN, MAX, MIN As Integer

Private Sub CMD_1_Click(Index As Integer)
If RAN = Index + 1 Then
MSG = MsgBox("the Bomb is {" & Index + 1 & "} PLAY AGAIN!", vbQuestion + vbYesNo, "BOMB")
    If MSG = vbYes Then
    Unload Me
    frmgame.Show
    Else
    End
    End If
Else
If Index + 1 < RAN Then
MIN = Index + 1
For X = 0 To Index
CMD_1(X).Visible = False
LBL_1.Caption = MIN & "~~" & MAX
Next
Else
MAX = Index + 1
For X = Index To 98
CMD_1(X).Visible = False
LBL_1.Caption = MIN & "~~" & MAX
Next
End If
End If
End Sub

Private Sub cmdback_Click()
frmmain.Visible = True
frmgame.Visible = False

End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub Form_Load()
Randomize
RAN = Int(Rnd * 100)
If Val(RAN) > 99 Then
RAN = 99
Else
If Val(RAN) < 1 Then
RAN = 1
End If
End If
    For I = 1 To 98
    DoEvents
    A = I Mod 10
    B = (I \ 10)
    Load CMD_1(I)
    CMD_1(I).Top = CMD_1(I).Top + (CMD_1(I).Height + 200) * B
    CMD_1(I).Left = CMD_1(I).Left + (CMD_1(I).Width + 200) * A
    CMD_1(I).Caption = I + 1
    CMD_1(I).Visible = True
    Next I
    MAX = 99
    MIN = 1
    LBL_1.Caption = MIN & "~~" & MAX
    

    

End Sub

