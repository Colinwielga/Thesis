VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C00000&
   Caption         =   "Simpsons "
   ClientHeight    =   8370
   ClientLeft      =   -135
   ClientTop       =   60
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   Picture         =   "main.frx":0000
   ScaleHeight     =   8370
   ScaleWidth      =   11490
   Begin VB.CommandButton cmdCharacters 
      Caption         =   "View Family"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   960
      TabIndex        =   6
      Top             =   2760
      Width           =   1935
   End
   Begin VB.PictureBox picSimp 
      Height          =   1215
      Left            =   3960
      Picture         =   "main.frx":8274
      ScaleHeight     =   1155
      ScaleWidth      =   3780
      TabIndex        =   5
      Top             =   360
      Width           =   3840
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H80000007&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   8520
      MaskColor       =   &H00800000&
      TabIndex        =   4
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton cmdHighscores 
      BackColor       =   &H80000007&
      Caption         =   "View Hall of Fame"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   8520
      MaskColor       =   &H00800000&
      TabIndex        =   3
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdhighscore 
      Caption         =   "View High Score"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   6480
      TabIndex        =   2
      Top             =   11760
      Width           =   2055
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Take the Test"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   4680
      Width           =   2055
   End
   Begin VB.PictureBox picSimp2 
      Height          =   5175
      Left            =   3480
      Picture         =   "main.frx":177D2
      ScaleHeight     =   5115
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   2520
      Width           =   4815
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'simpsons tv show test (final.vbp)
'main form (main.frm)
'Jim Berg
'October 30, 2005
'This is the main form and has all of the command buttons that will connect to other forms

Private Sub cmdCharacters_Click()
    frmmain.Hide
    frmCharacters.Show

End Sub

Private Sub cmdHighscores_Click(Index As Integer)
    frmhighscore.Show
    frmmain.Hide
End Sub


Private Sub cmdQuit_Click(Index As Integer)
End
End Sub


Private Sub cmdTest_Click(Index As Integer)
    frmSimpsons.Show
    frmmain.Hide
End Sub

