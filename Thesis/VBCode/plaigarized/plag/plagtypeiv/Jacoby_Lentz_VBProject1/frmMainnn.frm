VERSION 5.00
Begin VB.Form frmMain
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPurchase
      Caption         =   "Buy Entourage"
      Height          =   735
      Left            =   11640
      TabIndex        =   7
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton cmdSeasonOverviews
      Caption         =   "Season Overviews"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   9840
      Width           =   2175
   End
   Begin VB.CommandButton cmdWorksCited
      Caption         =   "Works Cited"
      Height          =   375
      Left            =   11760
      TabIndex        =   5
      Top             =   10800
      Width           =   2055
   End
   Begin VB.CommandButton cmdConvert
      Caption         =   "Convert Dollars to Pesos"
      Height          =   735
      Left            =   9720
      TabIndex        =   4
      Top             =   9840
      Width           =   2175
   End
   Begin VB.CommandButton cmdMovies
      Caption         =   "Movies"
      Height          =   735
      Left            =   7320
      TabIndex        =   3
      Top             =   9840
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit
      Caption         =   "Quit"
      Height          =   735
      Left            =   12120
      TabIndex        =   2
      Top             =   9840
      Width           =   2175
   End
   Begin VB.CommandButton cmdTrivia
      Caption         =   "Entourage Trivia"
      Height          =   735
      Left            =   4920
      TabIndex        =   1
      Top             =   9840
      Width           =   2175
   End
   Begin VB.CommandButton cmdMeetCharacters
      Caption         =   "Meet The Characters"
      Height          =   735
      Left            =   2520
      TabIndex        =   0
      Top             =   9840
      Width           =   2175
   End
   Begin VB.OLE OLE1
      BackColor       =   &H80000006&
      Class           =   "Package"
      DisplayType     =   1  'Icon
      Height          =   615
      Left            =   240
      OleObjectBlob   =   "frmMain.frx":240042
      SourceDoc       =   "M:\CS130\Jacoby_Lentz_VBProject1\entourage.mp3"
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMeetCharacters_Click()
    frmMain.Hide
    frmSeasonsOverview.Hide
    frmTrivia.Hide
    frmMovies.Hide
    frmConvert.Hide
    frmWorksCited.Hide
    frmAriGold.Hide
    frmEricMurphy.Hide
    frmJohnnyDrama.Hide
    frmTurtle.Hide
    frmVincentChase.Hide
    frmCharacters.Show
End Sub

Private Sub cmdMovies_Click()
    frmMovies.Show
    frmMain.Hide
    frmSeasonsOverview.Hide
    frmCharacters.Hide
    frmTrivia.Hide
    frmConvert.Hide
    frmWorksCited.Hide
    frmAriGold.Hide
    frmEricMurphy.Hide
    frmJohnnyDrama.Hide
    frmTurtle.Hide
    frmVincentChase.Hide
End Sub

Private Sub cmdConvert_Click()
    frmConvert.Show
    frmMain.Hide
    frmSeasonsOverview.Hide
    frmCharacters.Hide
    frmTrivia.Hide
    frmMovies.Hide
    frmWorksCited.Hide
    frmAriGold.Hide
    frmEricMurphy.Hide
    frmJohnnyDrama.Hide
    frmTurtle.Hide
    frmVincentChase.Hide
End Sub

Private Sub cmdPurchase_Click()
    frmCharacters.Hide
    frmTrivia.Hide
    frmMovies.Hide
    frmConvert.Hide
    frmWorksCited.Hide
    frmAriGold.Hide
    frmEricMurphy.Hide
    frmJohnnyDrama.Hide
    frmTurtle.Hide
    frmVincentChase.Hide
    frmMain.Hide

    frmPurchase.Show
End Sub

Private Sub cmdQuit_Click()
    MsgBox "You Can Have It All"
    End
End Sub

Private Sub cmdSeasonOverviews_Click()
    frmMain.Hide
    frmCharacters.Hide
    frmTrivia.Hide
    frmMovies.Hide
    frmConvert.Hide
    frmWorksCited.Hide
    frmAriGold.Hide
    frmEricMurphy.Hide
    frmJohnnyDrama.Hide
    frmTurtle.Hide
    frmVincentChase.Hide
    frmSeasonsOverview.Show
End Sub

Private Sub cmdTrivia_Click()
    frmMain.Hide
    frmSeasonsOverview.Hide
    frmCharacters.Hide
    frmMovies.Hide
    frmConvert.Hide
    frmWorksCited.Hide
    frmAriGold.Hide
    frmEricMurphy.Hide
    frmJohnnyDrama.Hide
    frmTurtle.Hide
    frmVincentChase.Hide
    frmTrivia.Show
End Sub

Private Sub cmdWorksCited_Click()
    frmMain.Hide
    frmSeasonsOverview.Hide
    frmCharacters.Hide
    frmTrivia.Hide
    frmMovies.Hide
    frmConvert.Hide
    frmAriGold.Hide
    frmEricMurphy.Hide
    frmJohnnyDrama.Hide
    frmTurtle.Hide
    frmVincentChase.Hide
    frmWorksCited.Show
End Sub
