VERSION 5.00
Begin VB.Form frmLoad 
   BackColor       =   &H0000C000&
   Caption         =   "Load"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Load"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   4560
      Picture         =   "Load.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      Height          =   735
      Left            =   120
      MaskColor       =   &H8000000F&
      TabIndex        =   3
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdNextForm 
      Caption         =   "Calculate Team Statistics"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0080C0FF&
      Height          =   3975
      Left            =   1800
      ScaleHeight     =   3915
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   2160
      Width           =   4335
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Stats"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblExplanation 
      BackColor       =   &H00FFFF80&
      Caption         =   "This program was designed to use the statistics of a local baseball team and determine the team leader in each category."
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdLoad_Click()
    'Ben Tollefson
    'March 22, 2006
    'This is to load the stats into the program
    Dim Pos As Integer
    Open App.Path & "\baseball.txt" For Input As #1
    Pos = 0
        picResults.Print "Name"; Tab(20); "Home Runs"; Tab(35); "Hits"; Tab(45); "At Bats"
        picResults.Print
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, Names(Pos), HR(Pos), Hits(Pos), AB(Pos)
        picResults.Print Names(Pos); Tab(20); HR(Pos); Tab(35); Hits(Pos); Tab(45); AB(Pos)
    Loop
    Size = Pos
End Sub

Private Sub cmdNextForm_Click()
'Ben Tollefson
'March 22, 2006
'This Button goes to the next form
    frmStats.Show
    frmLoad.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub
