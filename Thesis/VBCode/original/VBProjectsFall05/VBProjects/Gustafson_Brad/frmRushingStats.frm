VERSION 5.00
Begin VB.Form frmRushingStats 
   BackColor       =   &H00000000&
   Caption         =   "RushingStats"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPlayerRushing 
      Height          =   2175
      Left            =   3000
      Picture         =   "frmRushingStats.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   3120
      Width           =   4335
   End
   Begin VB.PictureBox picRushingStats 
      BackColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   240
      ScaleHeight     =   2835
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   120
      Width           =   7695
   End
   Begin VB.Label lblBarber 
      Caption         =   "        #24 Marion Barber"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   2655
   End
End
Attribute VB_Name = "frmRushingStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Player(1 To 10) As String
Dim Carries(1 To 10) As Double, Yards(1 To 10) As Double, Average(1 To 10) As Double, Longest(1 To 10) As Double, TouchDowns(1 To 10) As Double
Dim TempPlayer As String
Dim TempCarries As Double, TempYards As Double, TempAverage As Double, TempLongest As Double, TempTouchdowns As Double
Dim Pass As Integer
Dim I As Integer

Private Sub picRushingStats_Paint()
    Open App.Path & "\rushingstats.txt" For Input As #4
    picRushingStats.Print "Player"; Tab(20); "Carries"; Tab(30); "Yards", "Average", "Longest", "TouchDown"
    picRushingStats.Print "========================================================================"
    For I = 1 To 10
        Input #4, Player(I), Carries(I), Yards(I), Average(I), Longest(I), TouchDowns(I)
        picRushingStats.Print Player(I); Tab(20); Carries(I); Tab(30); Yards(I), Average(I), Longest(I), TouchDowns(I)
    Next I
    Close #4
End Sub


