VERSION 5.00
Begin VB.Form frmDecember 
   BackColor       =   &H000000FF&
   Caption         =   "December Schdule"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10155
   LinkTopic       =   "Form2"
   ScaleHeight     =   7665
   ScaleWidth      =   10155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoJanuary 
      Caption         =   "Go to January Schedule"
      Height          =   1335
      Left            =   3120
      TabIndex        =   3
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton cmdMainDecember 
      Caption         =   "Back to main page"
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   5880
      Width           =   1935
   End
   Begin VB.PictureBox picDecember 
      Height          =   4455
      Left            =   240
      ScaleHeight     =   4395
      ScaleWidth      =   4755
      TabIndex        =   1
      Top             =   1080
      Width           =   4815
   End
   Begin VB.CommandButton cmdDecemberGames 
      Caption         =   "Show me the games played in December."
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   8325
      Left            =   5400
      Picture         =   "frmDecember.frx":0000
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmDecember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDecemberGames_Click()
Dim Schedule(1 To 100) As String, GameNight(1 To 100) As String
Dim Pass As Integer, Pos As Integer, CTR As Single
Dim December(1 To 100) As Date
Open App.Path & "\schedule.txt" For Input As #1

CTR = 0
picDecember.Print "Opponent"; Tab(40); "Date" 'Prints heading
picDecember.Print "*************************************************************"
Do Until EOF(1) 'enters all data from array
    CTR = CTR + 1
    Input #1, Schedule(CTR), GameNight(CTR)
Loop
Close #1
    
For Pos = 1 To CTR
    If Left(GameNight(Pos), 2) = 12 Then 'takes the first 2 digits from month
    picDecember.Print Schedule(Pos); Tab(40); GameNight(Pos)
    End If 'this was tough to figure out
Next Pos 'the string function makes it possible to read dates in array
        
        

End Sub

Private Sub cmdGoJanuary_Click()
frmDecember.Hide
frmJanuary.Show
End Sub

Private Sub cmdMainDecember_Click()
frmDecember.Hide
frmStormBball.Show
End Sub
