VERSION 5.00
Begin VB.Form frmUsername 
   BackColor       =   &H00000000&
   Caption         =   "Username"
   ClientHeight    =   9855
   ClientLeft      =   4140
   ClientTop       =   840
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmUsername.frx":0000
   ScaleHeight     =   9855
   ScaleWidth      =   7425
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6000
      TabIndex        =   5
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue to the game"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   4
      Top             =   8880
      Width           =   2415
   End
   Begin VB.PictureBox picUsername 
      BackColor       =   &H00800080&
      Height          =   2295
      Left            =   1200
      ScaleHeight     =   2235
      ScaleWidth      =   4755
      TabIndex        =   3
      Top             =   5880
      Width           =   4815
   End
   Begin VB.CommandButton cmdUsername 
      Caption         =   "Calculate Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   8880
      Width           =   2535
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   3840
      Width           =   6735
   End
   Begin VB.Label lblName 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter your first and last name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1395
      Left            =   1440
      TabIndex        =   0
      Top             =   0
      Width           =   4440
   End
End
Attribute VB_Name = "frmUsername"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project:Deal or No Deal
'frmUsername
'Holly Reinking and Danielle Karp
'Written 3/22/09
'Purpose: To find a username for the player

Private Sub cmdContinue_Click()
frmUsername.Hide        'After the Username is found the player can continue on to the game
frmdealornodeal.Show

End Sub

Private Sub cmdQuit_Click() 'Used to quit the game
    End
End Sub

Private Sub cmdUsername_Click()         'In order to find the ID for the player, a combination of First and Last Name
Dim whoName As String
Dim h As Integer
Dim First As String, Last As String

whoName = txtName.Text
h = InStr(whoName, " ")
First = Left(whoName, h - 1)
Last = Right(whoName, Len(whoName) - (h + 2))
id = Left(First, 1) & Left(Last, 6)
picUsername.Print " Your player name is: "; Tab(5); id
cmdContinue.Enabled = True
cmdUsername.Enabled = False

End Sub



