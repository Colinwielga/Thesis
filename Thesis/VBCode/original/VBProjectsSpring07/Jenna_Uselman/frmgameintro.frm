VERSION 5.00
Begin VB.Form frmgameintro 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Welcome!"
   ClientHeight    =   4425
   ClientLeft      =   6075
   ClientTop       =   3270
   ClientWidth     =   8175
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   8175
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   3
      Top             =   1560
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000009&
      Caption         =   "return to menu"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   3375
   End
   Begin VB.CommandButton cmdplay 
      BackColor       =   &H80000009&
      Caption         =   "let's play!"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "enter name:"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label lblgameintro 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "welcome to abc show trivia!"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   7215
   End
End
Attribute VB_Name = "frmgameintro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Purpose of form: This form is the introduction to the trivia game. It allows the user to change
'                 their mind before the game starts, or prepares them for the game. Also allows
'                 user to enter their name.




Private Sub cmdplay_Click() 'This command button saves the user name and opens the trivia game form.
frmgame.Show
frmgameintro.Hide
uname = txtname.Text

End Sub

Private Sub Command2_Click() 'This command button takes the user back to the menu form.
frmMain.Show
frmgameintro.Hide
End Sub

