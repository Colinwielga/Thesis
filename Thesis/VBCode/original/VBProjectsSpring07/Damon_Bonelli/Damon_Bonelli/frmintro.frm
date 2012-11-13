VERSION 5.00
Begin VB.Form frmintro 
   BackColor       =   &H000000FF&
   Caption         =   "Johnnie Volleyball"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgotostats 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Look up Statistics"
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdgotofind 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Find a Player"
      Height          =   975
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdleave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      Height          =   975
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdload 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Load Players"
      Height          =   975
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Image imgintro 
      Height          =   4320
      Left            =   240
      Picture         =   "frmintro.frx":0000
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   6120
   End
   Begin VB.Label lblbesureto 
      Caption         =   "Be sure to click the load players button before you enter other parts of the program."
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   6120
      Width           =   6135
   End
End
Attribute VB_Name = "frmintro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdgotofind_Click()
'This button causes the program to switch
'to the find player screen
frmfind.Show
frmintro.Hide
End Sub

Private Sub cmdgotostats_Click()
'this button causes the program to switch
'to the statistics screen
frmStats.Show
frmintro.Hide
End Sub

Private Sub cmdleave_Click()
'quits the program
End
End Sub

Private Sub cmdLoad_Click()
'opens a file in the folder that contains all statistics concerning the players
'such as their names, hitting attempts, hitting errors, kills, blocks
'service aces, service errors, the number of games that they have played
'their numbers and their positions
Open App.Path & "\stats.txt" For Input As #1
Do Until EOF(1)
    ctr = ctr + 1
    Input #1, names(ctr), attempts(ctr), errors(ctr), kills(ctr), blocks(ctr), aces(ctr), se(ctr), games(ctr), jersey(ctr), position(ctr)
Loop
Close #1
End Sub

