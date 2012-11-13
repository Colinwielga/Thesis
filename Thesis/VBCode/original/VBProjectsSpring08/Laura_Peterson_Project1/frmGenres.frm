VERSION 5.00
Begin VB.Form frmGenres 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   10740
   ClientLeft      =   855
   ClientTop       =   645
   ClientWidth     =   13290
   LinkTopic       =   "Form1"
   ScaleHeight     =   10740
   ScaleWidth      =   13290
   Begin VB.CommandButton cmdFullList 
      BackColor       =   &H00808000&
      Caption         =   "See All Films"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7560
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H8000000C&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9360
      Width           =   6735
   End
   Begin VB.CommandButton cmdTriviaGame 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Play Laura's Movie Trivia Game!"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6360
      Width           =   2775
   End
   Begin VB.CommandButton cmdMisc 
      Height          =   2415
      Left            =   8400
      Picture         =   "frmGenres.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton cmdHorror 
      Height          =   3375
      Left            =   1440
      Picture         =   "frmGenres.frx":17C82
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton cmdAction 
      Height          =   4095
      Left            =   5040
      Picture         =   "frmGenres.frx":1A14F
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   2535
   End
   Begin VB.CommandButton cmdMusicals 
      Height          =   3255
      Left            =   8520
      Picture         =   "frmGenres.frx":1D75D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton cmdDrama 
      Height          =   3015
      Left            =   840
      Picture         =   "frmGenres.frx":20239
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblGenre 
      BackColor       =   &H80000012&
      Caption         =   "Pick A Genre"
      BeginProperty Font 
         Name            =   "Rage Italic"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   1215
      Left            =   4080
      TabIndex        =   0
      Top             =   600
      Width           =   5055
   End
End
Attribute VB_Name = "frmGenres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAction_Click()
frmAction.Show
End Sub

Private Sub cmdDrama_Click()
frmDrama.Show
End Sub

Private Sub cmdFullList_Click()
frmFullList.Show
End Sub

Private Sub cmdHorror_Click()
frmHorror.Show
End Sub

Private Sub cmdMisc_Click()
frmMisc.Show
End Sub

Private Sub cmdMusicals_Click()
frmMusicals.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdTriviaGame_Click()
frmGames.Show
End Sub
