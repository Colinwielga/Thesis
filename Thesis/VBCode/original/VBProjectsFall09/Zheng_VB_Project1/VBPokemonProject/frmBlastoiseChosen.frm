VERSION 5.00
Begin VB.Form frmBlastoiseChosen 
   Caption         =   "Blastoise"
   ClientHeight    =   10230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15345
   LinkTopic       =   "Form1"
   Picture         =   "frmBlastoiseChosen.frx":0000
   ScaleHeight     =   10230
   ScaleWidth      =   15345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   13920
      TabIndex        =   6
      Top             =   9480
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   735
      Left            =   13560
      TabIndex        =   5
      Top             =   8400
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
      Top             =   8280
      Width           =   2295
   End
   Begin VB.TextBox txtOpponent 
      Height          =   855
      Left            =   6600
      TabIndex        =   3
      Top             =   7200
      Width           =   1935
   End
   Begin VB.PictureBox picBlastosie 
      FillColor       =   &H00FFFFFF&
      Height          =   4935
      Left            =   5040
      Picture         =   "frmBlastoiseChosen.frx":6225
      ScaleHeight     =   4875
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   1680
      Width           =   5055
   End
   Begin VB.Label lblChoose 
      BackStyle       =   0  'Transparent
      Caption         =   "        Choose Your Opponent:                     Charizard or Venusaur"
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   6360
      TabIndex        =   2
      Top             =   6720
      Width           =   2535
   End
   Begin VB.Label lblBlastoise 
      BackStyle       =   0  'Transparent
      Caption         =   "You Have Chosen Blastoise!"
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
Attribute VB_Name = "frmBlastoiseChosen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Pokemon Battle
'frmBlastoiseChosen
'Eugene Zheng
'10/15/2009
'This form presents the Pokemon "Blastoise" to the user. From here, the user can choose which Pokemon to battle using Blastoise.


Private Sub cmdBack_Click()
'The Back button goes back to the main screen where the user chooses the pokemon to use.
'By using the hide function this can be accomplished.
frmBlastoiseChosen.Hide
frmChoosing.Show
End Sub

Private Sub cmdBattle_Click()

'The user chooses the his/her opponent using a text box. Thus "Opponent" needs to be dimed as a String function
'Depending on the input, a specific form will appear
Dim Opponent As String
Opponent = txtOpponent.Text

If Opponent = "Charizard" Then
    frmBlastoiseVsCharizard.Show
ElseIf Opponent = "Venusaur" Then
    frmBlastoiseVsVenusaur.Show
Else
MsgBox "Invalid Input: Check for spelling errors and make sure to capitalize the name of the opponent", , "Try Again"


End If

End Sub

Private Sub cmdQuit_Click()
'A simple quit button to exit the program
End
End Sub

