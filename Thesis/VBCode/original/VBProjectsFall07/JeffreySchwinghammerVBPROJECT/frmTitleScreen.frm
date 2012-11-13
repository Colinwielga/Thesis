VERSION 5.00
Begin VB.Form frmTitleScreen 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   8325
   ClientLeft      =   1350
   ClientTop       =   2685
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   10785
   Begin VB.PictureBox picTure 
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   720
      Picture         =   "frmTitleScreen.frx":0000
      ScaleHeight     =   4455
      ScaleWidth      =   9015
      TabIndex        =   13
      Top             =   240
      Width           =   9015
   End
   Begin VB.CommandButton cmdCheattoOutside 
      Caption         =   "Cheat to Outside"
      Height          =   735
      Left            =   12360
      TabIndex        =   12
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdcheattopostbattle 
      Caption         =   "Cheat to Post Battle"
      Height          =   495
      Left            =   14040
      TabIndex        =   11
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdCheattoBattle 
      Caption         =   "Cheat to Battle"
      Height          =   735
      Left            =   14280
      TabIndex        =   10
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdCheattoCom 
      Caption         =   "Cheat to computer"
      Height          =   615
      Left            =   14040
      TabIndex        =   9
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdcheattoLab 
      Caption         =   "Cheat to Lab"
      Height          =   735
      Left            =   12240
      TabIndex        =   8
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdcheattobookcase 
      Caption         =   "Cheat to BookCase"
      Height          =   1215
      Left            =   14520
      TabIndex        =   7
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton cmdCheattoLibrary 
      Caption         =   "Cheat to library"
      Height          =   1095
      Left            =   12240
      TabIndex        =   6
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton cmdCheattoHub 
      Caption         =   "Cheat to Hub"
      Height          =   1215
      Left            =   12000
      TabIndex        =   5
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   3600
      TabIndex        =   2
      Top             =   6840
      Width           =   4095
   End
   Begin VB.CommandButton cndCompletionList 
      Caption         =   "View Completion List"
      Height          =   855
      Left            =   3600
      TabIndex        =   1
      Top             =   5880
      Width           =   4095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Game"
      Height          =   855
      Left            =   3600
      TabIndex        =   0
      Top             =   4920
      Width           =   4095
   End
   Begin VB.Label lblSchwingStudios 
      Alignment       =   2  'Center
      Caption         =   "A Schwing Studios Production"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   4
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4320
      TabIndex        =   3
      Top             =   1440
      Width           =   2655
   End
End
Attribute VB_Name = "frmTitleScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheattoBattle_Click()
FrmBattle.Show
frmTitleScreen.Hide
End Sub

Private Sub cmdcheattobookcase_Click()
frmBookCase.Show
frmTitleScreen.Hide
End Sub

Private Sub cmdCheattoCom_Click()
frmComputer.Show
frmTitleScreen.Hide
End Sub

Private Sub cmdCheattoHub_Click()
frmHub.Show               'Changes Form, Starts the Game
frmTitleScreen.Hide
End Sub

Private Sub cmdcheattoLab_Click()
frmLab.Show
frmTitleScreen.Hide
End Sub

Private Sub cmdCheattoLibrary_Click()
frmLibrary.Show
frmTitleScreen.Hide
End Sub

Private Sub cmdCheattoOutside_Click()
frmTitleScreen.Hide
frmOutside.Show
End Sub

Private Sub cmdcheattopostbattle_Click()
frmBattleConclusion.Show
frmTitleScreen.Hide

End Sub

Private Sub cmdQuit_Click()
End         'Ends Program
End Sub

Private Sub cmdStart_Click()
frmIntro.Show               'Changes Form, Starts the Game
frmTitleScreen.Hide
End Sub

Private Sub cndCompletionList_Click()
frmPlayerList.Show
frmTitleScreen.Hide
End Sub

