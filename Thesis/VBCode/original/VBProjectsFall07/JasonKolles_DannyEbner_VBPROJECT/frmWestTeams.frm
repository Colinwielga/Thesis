VERSION 5.00
Begin VB.Form frmWestTeams 
   BackColor       =   &H00FF0000&
   Caption         =   "Western Conference Teams"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   14415
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8040
      Width           =   1815
   End
   Begin VB.CommandButton cmdBackHome 
      BackColor       =   &H000000FF&
      Caption         =   "Click Here to Go Back Home"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8040
      Width           =   1695
   End
   Begin VB.CommandButton cmdWolves 
      Height          =   1695
      Left            =   480
      Picture         =   "frmWestTeams.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdGrizzlies 
      Height          =   1575
      Index           =   13
      Left            =   12360
      Picture         =   "frmWestTeams.frx":0DD6
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdClippers 
      Height          =   1575
      Index           =   12
      Left            =   9720
      Picture         =   "frmWestTeams.frx":1900
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdLakers 
      Height          =   1575
      Index           =   11
      Left            =   6720
      Picture         =   "frmWestTeams.frx":2607
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdBlazers 
      Height          =   1575
      Index           =   10
      Left            =   3480
      Picture         =   "frmWestTeams.frx":30A2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdHornets 
      Height          =   1695
      Index           =   9
      Left            =   12360
      Picture         =   "frmWestTeams.frx":3CB6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdNuggets 
      Height          =   1695
      Index           =   8
      Left            =   9720
      Picture         =   "frmWestTeams.frx":4AB2
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdSpurs 
      Height          =   1695
      Index           =   7
      Left            =   6720
      Picture         =   "frmWestTeams.frx":51DC
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdSuns 
      Height          =   1695
      Index           =   6
      Left            =   3480
      Picture         =   "frmWestTeams.frx":5F89
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdWarriors 
      Height          =   1575
      Index           =   5
      Left            =   480
      Picture         =   "frmWestTeams.frx":6E91
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdJazz 
      Height          =   1695
      Index           =   4
      Left            =   480
      Picture         =   "frmWestTeams.frx":7D1E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdKings 
      Height          =   1695
      Index           =   3
      Left            =   12360
      Picture         =   "frmWestTeams.frx":85B0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdSonics 
      Height          =   1695
      Index           =   2
      Left            =   9720
      Picture         =   "frmWestTeams.frx":9041
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdMavs 
      Height          =   1695
      Index           =   1
      Left            =   6720
      Picture         =   "frmWestTeams.frx":98BE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdRockets 
      Height          =   1695
      Index           =   0
      Left            =   3480
      Picture         =   "frmWestTeams.frx":A8E1
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmWestTeams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBackHome_Click()
frmHome.Show
frmWestTeams.Hide

End Sub

Private Sub cmdBlazers_Click(Index As Integer)
frmBlazers.Show
frmWestTeams.Hide

End Sub

Private Sub cmdClippers_Click(Index As Integer)
    'bring user to Clippers page
    
frmClippers.Show
frmWestTeams.Hide

End Sub

Private Sub cmdGrizzlies_Click(Index As Integer)
    'bring user to Grizzlies page
    
frmGrizzlies.Show
frmWestTeams.Hide
End Sub

Private Sub cmdHornets_Click(Index As Integer)
    'brings user to Hornets page
    
frmHornets.Show
frmWestTeams.Hide

End Sub

Private Sub cmdJazz_Click(Index As Integer)
    'bring user to Jazz page
frmJazz.Show
frmWestTeams.Hide

End Sub

Private Sub cmdKings_Click(Index As Integer)
    'bring user to kings page
    
frmKings.Show
frmWestTeams.Hide

End Sub

Private Sub cmdLakers_Click(Index As Integer)
    'bring user to lakers page
    
frmLakers.Show
frmWestTeams.Hide

End Sub

Private Sub cmdMavs_Click(Index As Integer)
    'bring user to mavs page
    
frmMavericks.Show
frmWestTeams.Hide

End Sub

Private Sub cmdNuggets_Click(Index As Integer)
    'bring user to nuggs page
    
frmNuggets.Show
frmWestTeams.Hide

End Sub

Private Sub cmdQuit_Click()
End

End Sub

Private Sub cmdRockets_Click(Index As Integer)
    'bring user to rockets page
    
frmRockets.Show
frmWestTeams.Hide

End Sub

Private Sub cmdSonics_Click(Index As Integer)
    'bring user to Sonics page
    
frmSonics.Show
frmWestTeams.Hide

End Sub

Private Sub cmdSpurs_Click(Index As Integer)
    'bring user to spurs page
    
frmSpurs.Show
frmWestTeams.Hide

End Sub

Private Sub cmdSuns_Click(Index As Integer)
    'bring user to Suns page
    
frmSuns.Show
frmWestTeams.Hide

End Sub

Private Sub cmdWarriors_Click(Index As Integer)
    'bring user to Warriors page
    
frmWarriors.Show
frmWestTeams.Hide

End Sub

Private Sub cmdWolves_Click()
    'bring user to Wolves page
    
frmWolves.Show
frmWestTeams.Hide

End Sub
