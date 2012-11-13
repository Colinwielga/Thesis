VERSION 5.00
Begin VB.Form frmChampions 
   BackColor       =   &H00800000&
   Caption         =   "Champions League"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton cmddream 
      Caption         =   "Dream Team"
      Height          =   855
      Left            =   4800
      TabIndex        =   6
      Top             =   5520
      Width           =   2295
   End
   Begin VB.CommandButton cmdPlaces 
      Caption         =   "Venue"
      Height          =   855
      Left            =   1560
      TabIndex        =   5
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CommandButton cmdgoal 
      Caption         =   "Top Scorers"
      Height          =   975
      Left            =   6600
      TabIndex        =   4
      Top             =   4320
      Width           =   2415
   End
   Begin VB.CommandButton CmdSecond 
      Caption         =   "Runner Up"
      Height          =   975
      Left            =   3480
      TabIndex        =   3
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CommandButton CmdWinner 
      Caption         =   "CHAMPIONS"
      Height          =   975
      Left            =   600
      TabIndex        =   2
      Top             =   4320
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   2160
      Picture         =   "frmChampions.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   5235
      TabIndex        =   1
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label lblCreated 
      BackStyle       =   0  'Transparent
      Caption         =   "Created by: Gustavo Vallarino"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6840
      Width           =   3735
   End
End
Attribute VB_Name = "frmChampions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This is the main page of the project and here you can choose where you wan to be directed to
'funcions used are basically hide and show


Private Sub cmddream_Click()
frmDream.Show
frmChampions.Hide
End Sub

Private Sub cmdgoal_Click()
frmGoal.Show
frmChampions.Hide
End Sub

Private Sub cmdPlaces_Click()
frmVenues.Show
frmChampions.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub CmdSecond_Click()
frmRunnerUp.Show
frmChampions.Hide
End Sub

Private Sub CmdWinner_Click()
frmWinner.Show
frmChampions.Hide
End Sub

