VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main Menu"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit Driver Seat"
      Height          =   615
      Left            =   5640
      TabIndex        =   5
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton cmdCrew 
      Caption         =   "Visit Pit Row"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton cmdChase 
      Caption         =   "Race to the chase Sprint Cup 2009"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdTeam 
      Caption         =   "Meet Hendrick Motorsports and its drivers"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton cmdTracks 
      Caption         =   "Race Tracks in the Chase for the Sprint Cup Championship"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton cmdHistory 
      Caption         =   "History of Nascar"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   8910
      Left            =   0
      Picture         =   "frmMain.frx":0000
      Top             =   0
      Width           =   7560
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Introduction to NASCAR
'Form Main
'Colin Roberts and Luke Hommerding
'Written 10/18/09
'Purpose is to provide user with multiple options to experience the sport of NASCAR using
'Multiple forms
Option Explicit
'takes users to the chase form
Private Sub cmdChase_Click()
    frmChase.Show
    frmMain.Hide
End Sub
'takes user to the history form
Private Sub cmdHistory_Click()
    frmHistory.Show
    frmMain.Hide
End Sub
'takes user to the crew form
Private Sub cmdCrew_Click()
    frmCrew.Show
    frmMain.Hide
End Sub
'takes user to the team form
Private Sub cmdTeam_Click()
    frmTeam.Show
    frmMain.Hide
End Sub
'takes user to the tracks form
Private Sub cmdTracks_Click()
    frmTracks.Show
    frmMain.Hide
End Sub
'quits the whole program
Private Sub cmdQuit_Click()
End
End Sub
