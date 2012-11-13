VERSION 5.00
Begin VB.Form frmMarch 
   BackColor       =   &H0080FF80&
   Caption         =   "March Schedule"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9990
   LinkTopic       =   "Form5"
   ScaleHeight     =   7440
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBackMarch 
      Caption         =   "Back to main page"
      Height          =   1095
      Left            =   1320
      TabIndex        =   2
      Top             =   5520
      Width           =   1935
   End
   Begin VB.PictureBox picMarch 
      Height          =   3735
      Left            =   240
      ScaleHeight     =   3675
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   1320
      Width           =   4335
   End
   Begin VB.CommandButton cmdMarchSchedule 
      Caption         =   "Show me the games played in March."
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   6150
      Left            =   4920
      Picture         =   "frmMarch.frx":0000
      Top             =   480
      Width           =   4500
   End
End
Attribute VB_Name = "frmMarch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBackMarch_Click()
frmMarch.Hide
frmStormBball.Show
End Sub

Private Sub cmdMarchSchedule_Click()
Dim Schedule(1 To 100) As String, GameNight(1 To 100) As String
Dim Pass As Integer, Pos As Integer, CTR As Single
Dim March(1 To 100) As Date

Open App.Path & "\schedule.txt" For Input As #1

CTR = 0

picMarch.Print "Opponent"; Tab(40); "Date"
picMarch.Print "*************************************************************"
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Schedule(CTR), GameNight(CTR)
Loop
Close #1
    
For Pos = 1 To CTR
    If Left(GameNight(Pos), 2) = 3 Then
    picMarch.Print Schedule(Pos); Tab(40); GameNight(Pos)
    End If
Next Pos
End Sub
