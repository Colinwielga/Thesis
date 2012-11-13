VERSION 5.00
Begin VB.Form frm2005DraftPicks 
   BackColor       =   &H0000FFFF&
   Caption         =   "2005DraftPicks"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDraftPicks 
      BackColor       =   &H0080FFFF&
      Height          =   2055
      Left            =   240
      ScaleHeight     =   1995
      ScaleWidth      =   6315
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "frm2005DraftPicks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim Player(1 To 8) As String, Position(1 To 8) As String, College(1 To 8) As String
Dim Round(1 To 8) As Double, OverallPick(1 To 8) As Double

Private Sub picDraftPicks_Paint()
    Open App.Path & "\2005draftpicks.txt" For Input As #7
    picDraftPicks.Print "Player"; Tab(20); "Round #"; Tab(32); "Overall Draft Pick", "Position", "College"
    picDraftPicks.Print "********************************************************************************************************"
    For I = 1 To 8
        Input #7, Player(I), Round(I), OverallPick(I), Position(I), College(I)
        picDraftPicks.Print Player(I); Tab(22); Round(I); Tab(39); OverallPick(I), Position(I), College(I)
    Next I
    Close #7
End Sub
