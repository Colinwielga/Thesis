VERSION 5.00
Begin VB.Form frmJanuary 
   BackColor       =   &H00FFC0C0&
   Caption         =   "January Schedule"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9945
   LinkTopic       =   "Form3"
   ScaleHeight     =   7365
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoFebruary 
      BackColor       =   &H00FF8080&
      Caption         =   "Go to February Schedule"
      Height          =   1215
      Left            =   2640
      TabIndex        =   3
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton cmdMainJanuary 
      Caption         =   "Back to main page"
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   5640
      Width           =   1815
   End
   Begin VB.PictureBox picJanuary 
      Height          =   4095
      Left            =   240
      ScaleHeight     =   4035
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   1200
      Width           =   4335
   End
   Begin VB.CommandButton cmdJanuaryGames 
      Caption         =   "Show me the games played in January."
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   6750
      Left            =   4920
      Picture         =   "frmJanuary.frx":0000
      Top             =   240
      Width           =   4755
   End
End
Attribute VB_Name = "frmJanuary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdGoFebruary_Click()
frmJanuary.Hide
frmFebruary.Show
End Sub

Private Sub cmdJanuaryGames_Click()
Dim Schedule(1 To 100) As String, GameNight(1 To 100) As String
Dim Pass As Integer, Pos As Integer, CTR As Single
Dim January(1 To 100) As Date

Open App.Path & "\schedule.txt" For Input As #1

CTR = 0

picJanuary.Print "Opponent"; Tab(40); "Date"
picJanuary.Print "*************************************************************"
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Schedule(CTR), GameNight(CTR)
Loop
Close #1
    
For Pos = 1 To CTR
    If Left(GameNight(Pos), 2) = 1 Then
    picJanuary.Print Schedule(Pos); Tab(40); GameNight(Pos)
    End If
Next Pos
End Sub

Private Sub cmdMainJanuary_Click()
frmJanuary.Hide
frmStormBball.Show
End Sub
