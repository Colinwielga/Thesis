VERSION 5.00
Begin VB.Form frmFebruary 
   BackColor       =   &H00FF0000&
   Caption         =   "February Schedule"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10020
   LinkTopic       =   "Form4"
   ScaleHeight     =   7395
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoMarch 
      Caption         =   "Go to March Schedule"
      Height          =   975
      Left            =   2640
      TabIndex        =   3
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton cmdMainFeb 
      Caption         =   "Back to main page"
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   6000
      Width           =   1935
   End
   Begin VB.PictureBox picFebruary 
      Height          =   4335
      Left            =   360
      ScaleHeight     =   4275
      ScaleWidth      =   4155
      TabIndex        =   1
      Top             =   1320
      Width           =   4215
   End
   Begin VB.CommandButton cmdFebruarySchedule 
      Caption         =   "Show me the games played in February."
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
   Begin VB.Image Image2 
      Height          =   4500
      Left            =   5280
      Picture         =   "frmFebruary.frx":0000
      Top             =   3000
      Width           =   4500
   End
   Begin VB.Image Image1 
      Height          =   2445
      Left            =   5760
      Picture         =   "frmFebruary.frx":8FA1
      Top             =   480
      Width           =   3240
   End
End
Attribute VB_Name = "frmFebruary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFebruarySchedule_Click()
Dim Schedule(1 To 100) As String, GameNight(1 To 100) As String
Dim Pass As Integer, Pos As Integer, CTR As Single
Dim February(1 To 100) As Date

Open App.Path & "\schedule.txt" For Input As #1

CTR = 0

picFebruary.Print "Opponent"; Tab(40); "Date"
picFebruary.Print "*************************************************************"
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Schedule(CTR), GameNight(CTR)
Loop
Close #1
    
For Pos = 1 To CTR
    If Left(GameNight(Pos), 2) = 2 Then
    picFebruary.Print Schedule(Pos); Tab(40); GameNight(Pos)
    End If
Next Pos
End Sub

Private Sub cmdGoMarch_Click()
frmFebruary.Hide
frmMarch.Show
End Sub

Private Sub cmdMainFeb_Click()
frmStormBball.Show
frmFebruary.Hide
End Sub
