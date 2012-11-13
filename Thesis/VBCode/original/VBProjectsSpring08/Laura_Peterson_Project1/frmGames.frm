VERSION 5.00
Begin VB.Form frmGames 
   Caption         =   "Trivia Game"
   ClientHeight    =   6540
   ClientLeft      =   2025
   ClientTop       =   2985
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   Picture         =   "frmGames.frx":0000
   ScaleHeight     =   6540
   ScaleWidth      =   9420
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton cmdTriviaGame 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Play Laura's Movie Trivia Game"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   3735
   End
   Begin VB.CommandButton cmdPictureGame 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Play Laura's Picture Trivia Game"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   3735
   End
End
Attribute VB_Name = "frmGames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Laura's Movie Gallery
'frmGames
'Laura Peterson
'3/14/2008
'This form will take the user to whichever game they choose as well as the main menu
Private Sub cmdPictureGame_Click()
frmPictureGame.Show
End Sub

Private Sub cmdReturn_Click()
frmGenres.Show
frmGames.Hide
End Sub

Private Sub cmdTriviaGame_Click()
frm1TriviaGame.Show
End Sub
