VERSION 5.00
Begin VB.Form frmCharizardChosen 
   Caption         =   "Charizard"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14835
   LinkTopic       =   "Form2"
   Picture         =   "frmCharizardChosen.frx":0000
   ScaleHeight     =   9420
   ScaleWidth      =   14835
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOpponent 
      Height          =   855
      Left            =   6600
      TabIndex        =   5
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   735
      Left            =   13200
      TabIndex        =   4
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   13440
      TabIndex        =   3
      Top             =   7920
      Width           =   1095
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
      TabIndex        =   2
      Top             =   7440
      Width           =   2175
   End
   Begin VB.PictureBox picCharizard 
      Height          =   4215
      Left            =   4800
      Picture         =   "frmCharizardChosen.frx":6225
      ScaleHeight     =   4155
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   1680
      Width           =   5415
   End
   Begin VB.Label lblChoose 
      BackStyle       =   0  'Transparent
      Caption         =   "        Choose Your Opponent:                     Blastoise or Venusaur"
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   6360
      TabIndex        =   6
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Label lblCharizard 
      BackStyle       =   0  'Transparent
      Caption         =   "You Have Chosen Charizard!"
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
      Left            =   5160
      TabIndex        =   1
      Top             =   1080
      Width           =   4575
   End
End
Attribute VB_Name = "frmCharizardChosen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Pokemon Project
'frmCharizardChosen
'Eugene Zheng
'10/10/2009
'Using this screen, the user decides which opponent to fight
'There are two outcomes, Blastoise and Venusaur
'The user tells the computer what to do by using the text boxes


Option Explicit

Private Sub cmdBack_Click()
'A return button
frmCharizardChosen.Hide
frmChoosing.Show
End Sub

Private Sub cmdBattle_Click()

'This is the code to determine which pokemon to fight
Dim Opponent As String
Opponent = txtOpponent.Text

If Opponent = "Blastoise" Then
    frmCharizardVsBlastoise.Show
ElseIf Opponent = "Venusaur" Then
    frmCharizardVsVenusaur.Show
Else
MsgBox "Invalid Input: Check for spelling errors and make sure to capitalize the name of the opponent", , "Try Again"


End If

End Sub

Private Sub cmdQuit_Click()
'Simple Quit button
End
End Sub

Private Sub Form_Load()
'Change font color
lblChoose.ForeColor = vbRed
End Sub

