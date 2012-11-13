VERSION 5.00
Begin VB.Form frmHome 
   BackColor       =   &H00000000&
   Caption         =   "Fantasy NBA Agent"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16050
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10155
   ScaleWidth      =   16050
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      Caption         =   "Exit from Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   12000
      TabIndex        =   6
      Top             =   7800
      Width           =   3015
      Begin VB.CommandButton cmdQuit 
         BackColor       =   &H00FF0000&
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdPictureForm 
      BackColor       =   &H000000FF&
      Caption         =   "Search for Player Pictures"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton cmdAllWest 
      BackColor       =   &H000000FF&
      Caption         =   "Go To All Western Conference Starters' Stats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton cmdFindPlayers 
      BackColor       =   &H000000FF&
      Caption         =   "Search for Western Conference Starters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdFindTeam 
      BackColor       =   &H000000FF&
      Caption         =   "Find Western Conference Team"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   5175
      Left            =   3840
      Picture         =   "NBA_Project.frx":0000
      ScaleHeight     =   5115
      ScaleWidth      =   6915
      TabIndex        =   0
      Top             =   1920
      Width           =   6975
   End
   Begin VB.Label lblTitle_Home 
      BackColor       =   &H00000000&
      Caption         =   "Western Conference Fantasy NBA Agent"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1155
      Left            =   3840
      TabIndex        =   3
      Top             =   360
      Width           =   7095
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAllWest_Click()
    'brings user to the form that has all West players
frmShowAllWest.Show
frmHome.Hide

End Sub

Private Sub cmdFindPlayers_Click()
    'brings user to form that allows them to
    'search players in the West
frmSearchWestPlayers.Show
frmHome.Hide

End Sub

Private Sub cmdFindTeam_Click()
    'brings user to form that has all West teams
    
frmWestTeams.Show
frmHome.Hide

End Sub

Private Sub cmdPictureForm_Click()
    'brings user to form that allows
    'them to view pictures of select players
    
frmPictures.Show
frmHome.Hide

End Sub

Private Sub cmdQuit_Click()
End

End Sub
