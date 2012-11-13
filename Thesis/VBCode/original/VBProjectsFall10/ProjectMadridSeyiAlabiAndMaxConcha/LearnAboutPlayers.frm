VERSION 5.00
Begin VB.Form Statistics 
   Caption         =   "Statistics"
   ClientHeight    =   12675
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15390
   LinkTopic       =   "Form3"
   Picture         =   "LearnAboutPlayers.frx":0000
   ScaleHeight     =   12675
   ScaleWidth      =   15390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back To Information"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      TabIndex        =   11
      Top             =   0
      Width           =   3135
   End
   Begin VB.CommandButton cmdCarvalho 
      Caption         =   "Ricardo Carvalho# 2"
      Height          =   975
      Left            =   3360
      TabIndex        =   10
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdPepe 
      Caption         =   "Pepe #3"
      Height          =   975
      Left            =   3360
      TabIndex        =   9
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdRamos 
      Caption         =   "Sergio Ramos #4"
      Height          =   975
      Left            =   3360
      TabIndex        =   8
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmdMarcelo 
      Caption         =   "Marcelo #12"
      Height          =   975
      Left            =   3480
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdHiguain 
      Caption         =   "Gonzalo Higuain #20"
      Height          =   975
      Left            =   10200
      TabIndex        =   6
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdRonaldo 
      Caption         =   "Cristiano Ronaldo #7"
      Height          =   975
      Left            =   10200
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdDimaria 
      Caption         =   "Angel Di Maria #22"
      Height          =   975
      Left            =   8760
      TabIndex        =   4
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton cmdBenzema 
      Caption         =   "Karim Benzema  #9"
      Height          =   975
      Left            =   8640
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdAlonso 
      Caption         =   "Xabi Alonso #14"
      Height          =   975
      Left            =   6600
      TabIndex        =   2
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdOzil 
      Caption         =   "Mesut Ozil #23"
      Height          =   975
      Left            =   6600
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdCasillas 
      Caption         =   "Iker Casillas #1"
      Height          =   975
      Left            =   1320
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
End
Attribute VB_Name = "Statistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form allows the user to pick each player at their likeness. Each player when clicked,
'will show in another form the players' records, picture, and bio.


Option Explicit
Private Sub cmdAlonso_Click()
selectedPlayer = 14
Alonso.Show
Me.Hide
End Sub

Private Sub cmdBack_Click()
Information.Show
Statistics.Hide
PlayersStat.Hide
OpenPage.Hide
Me.Hide
End Sub

Private Sub cmdBenzema_Click()
selectedPlayer = 9
Benzema.Show
Me.Hide
End Sub

Private Sub cmdCarvalho_Click()
selectedPlayer = 2
Carvalho.Show
Me.Hide
End Sub

Private Sub cmdCasillas_Click()
selectedPlayer = 1
Casillas.Show
Me.Hide
End Sub

Private Sub cmdDimaria_Click()
selectedPlayer = 22
DiMaria.Show
Me.Hide
End Sub

Private Sub cmdHiguain_Click()
selectedPlayer = 20
Higuain.Show
Me.Hide
End Sub

Private Sub cmdMarcelo_Click()
selectedPlayer = 12
Marcelo.Show
Me.Hide
End Sub

Private Sub cmdOzil_Click()
selectedPlayer = 23
Ozil.Show
Me.Hide
End Sub

Private Sub cmdPepe_Click()
selectedPlayer = 3
Pepe.Show
Me.Hide
End Sub

Private Sub cmdRamos_Click()
selectedPlayer = 4
Ramos.Show
Me.Hide
End Sub

Private Sub cmdronaldo_Click()
selectedPlayer = 7
PlayersStat.Show
Me.Hide

End Sub
