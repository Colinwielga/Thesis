VERSION 5.00
Begin VB.Form frmSetup 
   Caption         =   "Johnyopoly"
   ClientHeight    =   6390
   ClientLeft      =   2100
   ClientTop       =   1845
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   11055
   Begin VB.CommandButton cmdok4 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7440
      TabIndex        =   14
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton cmdok3 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7440
      TabIndex        =   13
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton cmdok2 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7440
      TabIndex        =   12
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton cmdok1 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7440
      TabIndex        =   11
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit Names"
      Enabled         =   0   'False
      Height          =   1815
      Left            =   8040
      TabIndex        =   10
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox txtP4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   5280
      Width           =   3015
   End
   Begin VB.TextBox txtP3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   4800
      Width           =   3015
   End
   Begin VB.TextBox txtP2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   4320
      Width           =   3015
   End
   Begin VB.TextBox txtP1 
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Label Label6 
      Caption         =   "Player Four:"
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Player Three:"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Player Two:"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Player One:"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Please enter the names ot the players.  Two to four players can play this game."
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   3360
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Welcome to Johnyopoly!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2895
      Left            =   2640
      TabIndex        =   1
      Top             =   360
      Width           =   5535
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form starts with the first input position and will not allow a player to do anyting until a name is entered.
'when each name is entered the ok command next to it will become enabled, once clicked the next text box will become enabled.
'after two names are entered the player can continue entering names or click submit to return to the main frmBoard.
Private Sub cmdok1_Click()
cmdok1.Enabled = False
txtP1.Enabled = False
txtP2.Enabled = True
NOP = 1
End Sub

Private Sub cmdok2_Click()
cmdok1.Enabled = False
cmdok2.Enabled = False
txtP1.Enabled = False
txtP2.Enabled = False
txtP3.Enabled = True
cmdSubmit.Enabled = True
NOP = NOP + 1
End Sub

Private Sub cmdok3_Click()
cmdok1.Enabled = False
cmdok2.Enabled = False
cmdok3.Enabled = False
txtP1.Enabled = False
txtP2.Enabled = False
txtP3.Enabled = False
txtP4.Enabled = True
NOP = NOP + 1

End Sub

Private Sub cmdok4_Click()
cmdok1.Enabled = False
cmdok2.Enabled = False
cmdok3.Enabled = False
cmdok4.Enabled = False
txtP1.Enabled = False
txtP2.Enabled = False
txtP3.Enabled = False
txtP4.Enabled = False
NOP = NOP + 1

End Sub

Private Sub cmdSubmit_Click()
Player(1) = txtP1.Text
Player(2) = txtP2.Text
Player(3) = txtP3.Text
Player(4) = txtP4.Text
frmSetup.Visible = False
frmBoard.Visible = True
frmBoard.cmdLoad.Enabled = True
MsgBox Player(1) & ", Load the Property and roll the dice."



End Sub

Private Sub txtP1_Change()
If txtP1.Text <> "" Then
    cmdok1.Enabled = True
End If
End Sub

Private Sub txtP2_Change()
If txtP2.Text <> "" Then
    cmdok2.Enabled = True
End If
End Sub

Private Sub txtP3_Change()
If txtP3.Text <> "" Then
    cmdok3.Enabled = True
End If
End Sub

Private Sub txtP4_Change()
If txtP4.Text <> "" Then
    cmdok4.Enabled = True
End If
End Sub
