VERSION 5.00
Begin VB.Form frmVenusaurChosen 
   Caption         =   "Venusaur"
   ClientHeight    =   10245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   Picture         =   "frmVenusaurChosen.frx":0000
   ScaleHeight     =   10245
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   13800
      TabIndex        =   6
      Top             =   9480
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   735
      Left            =   13440
      TabIndex        =   5
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton cmdBattle 
      Caption         =   "Battle!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6480
      TabIndex        =   4
      Top             =   7800
      Width           =   2175
   End
   Begin VB.TextBox txtOpponent 
      Height          =   855
      Left            =   6600
      TabIndex        =   3
      Top             =   6720
      Width           =   1935
   End
   Begin VB.PictureBox picVenusaur 
      Height          =   4335
      Left            =   4920
      Picture         =   "frmVenusaurChosen.frx":6225
      ScaleHeight     =   4275
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   1680
      Width           =   5055
   End
   Begin VB.Label lblChoose 
      BackStyle       =   0  'Transparent
      Caption         =   "        Choose Your Opponent:                     Charizard or Blastoise"
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   6360
      TabIndex        =   2
      Top             =   6240
      Width           =   2535
   End
   Begin VB.Label lblVenusaur 
      BackStyle       =   0  'Transparent
      Caption         =   "You Have Chosen Venusaur!"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   1080
      Width           =   4575
   End
End
Attribute VB_Name = "frmVenusaurChosen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Pokemon Project
'FrmVenusaurChosen
'Eugene Zheng
'10/12/2009
'If this screen is brought up, then it means that the user has chosen Venusaur for the time being
'He/she can decide whether to fight Blastoise or Charizard by entering the name into the text box
Option Explicit

Private Sub cmdBack_Click()
'Return button
frmVenusaurChosen.Hide
frmChoosing.Show
End Sub

Private Sub cmdBattle_Click()

'This button determines the opponent
Dim Opponent As String
Opponent = txtOpponent.Text

If Opponent = "Charizard" Then
    'if its charizard
    frmVenusaurVsCharizard.Show
ElseIf Opponent = "Blastoise" Then
    'If its Blastoise
    frmVenusaurVsBlastoise.Show
Else
'Just in case there is a spelling error or input input
MsgBox "Invalid Input: Check for spelling errors and make sure to capitalize the name of the opponent", , "Try Again"
End If

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub Form_Load()
lblVenusaur.ForeColor = vbGreen
lblChoose.ForeColor = vbGreen
End Sub
