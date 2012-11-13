VERSION 5.00
Begin VB.Form frmIndividualStats 
   BackColor       =   &H00000000&
   Caption         =   "IndividualStats"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picIndividualStats 
      BackColor       =   &H00C0C000&
      Height          =   6735
      Left            =   120
      ScaleHeight     =   6675
      ScaleWidth      =   7275
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmIndividualStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim Player(1 To 32) As String
Dim Total(1 To 32) As Double, Tackles(1 To 32) As Double, Assists(1 To 32) As Double, Sacks(1 To 32) As Double, FumbleRec(1 To 32) As Double

Private Sub picIndividualStats_Paint()
    Open App.Path & "\individualdefense.txt" For Input As #5
    picIndividualStats.Print "Player"; Tab(20); "Total Tackles"; Tab(35); "Solo Tackles"; Tab(51); "Assisted Tackles"; Tab(71); "Sacks"; Tab(80); "Fumble Rec."
    picIndividualStats.Print "____________________________________________________________________________"
    For I = 1 To 32
        Input #5, Player(I), Total(I), Tackles(I), Assists(I), Sacks(I), FumbleRec(I)
        picIndividualStats.Print Player(I); Tab(24); Total(I); Tab(39); Tackles(I); Tab(57); Assists(I), Sacks(I), FumbleRec(I)
    Next I
    Close #5
End Sub
