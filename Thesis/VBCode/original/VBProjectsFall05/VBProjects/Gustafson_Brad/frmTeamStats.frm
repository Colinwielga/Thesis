VERSION 5.00
Begin VB.Form frmTeamStats 
   BackColor       =   &H00000080&
   Caption         =   "TeamStats"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCowboys 
      Height          =   2655
      Left            =   4560
      Picture         =   "frmTeamStats.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.PictureBox picTeamStats 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2595
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmTeamStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StatName(1 To 10) As String
Dim Cowboys(1 To 10) As Double, Opponents(1 To 10) As Double
Dim I As Integer

Private Sub picTeamStats_Paint()
    Open App.Path & "\teamstats.txt" For Input As #3
    picTeamStats.Print "Stat"; Tab(30); "Cowboys", "Opponents"
    picTeamStats.Print "-------------------------------------------------------------------------------------------"
    For I = 1 To 10
        Input #3, StatName(I), Cowboys(I), Opponents(I)
        picTeamStats.Print StatName(I); Tab(30); Cowboys(I), Opponents(I)
    Next I
    Close #3
    
End Sub

