VERSION 5.00
Begin VB.Form frmInterceptionStats 
   BackColor       =   &H00004000&
   Caption         =   "InterceptionStats"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picIntercept 
      Height          =   2655
      Left            =   240
      Picture         =   "frmInterceptionStats.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   6915
      TabIndex        =   1
      Top             =   2280
      Width           =   6975
   End
   Begin VB.PictureBox picInterceptionStats 
      Height          =   1815
      Left            =   240
      ScaleHeight     =   1755
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "frmInterceptionStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim Player(1 To 6) As String
Dim Number(1 To 6) As Double, Yards(1 To 6) As Double, Average(1 To 6) As Double, Longest(1 To 6) As Double, TouchDowns(1 To 6) As Double


Private Sub picInterceptionStats_Paint()
    Open App.Path & "\interceptions.txt" For Input As #6
    picInterceptionStats.Print "Player"; Tab(23); "Number"; Tab(33); "Yards", "Average", "Longest", "Touchdowns"
    picInterceptionStats.Print "_________________________________________________________________________"
    For I = 1 To 6
        Input #6, Player(I), Number(I), Yards(I), Average(I), Longest(I), TouchDowns(I)
        picInterceptionStats.Print Player(I); Tab(24); Number(I); Tab(34); Yards(I); Tab(44); Average(I); Tab(59); Longest(I); Tab(75); TouchDowns(I)
    Next I
    Close #6
End Sub
