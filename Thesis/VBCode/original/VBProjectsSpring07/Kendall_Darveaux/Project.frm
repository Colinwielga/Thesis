VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   1335
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox picresults 
      Height          =   2775
      Left            =   360
      ScaleHeight     =   2715
      ScaleWidth      =   5835
      TabIndex        =   2
      Top             =   1680
      Width           =   5895
   End
   Begin VB.CommandButton Cmdtotal 
      Caption         =   "Total Score"
      Height          =   1335
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Cmddive 
      Caption         =   "Dive Score"
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim total As Single, CTR As Integer


Private Sub Cmddive_Click()
Dim score1 As Single, score2 As Single, score3 As Single, dd As Single
Dim score4 As Single, score5 As Single, temp As Single
Dim sum As Single, divetotal As Single


dd = InputBox("Enter Degree of Difficulty", "DD")
score1 = InputBox("Enter Score 1", "Score")
score2 = InputBox("Enter Score 2", "Score")
score3 = InputBox("Enter Score 3", "Score")
score4 = InputBox("Enter Score 4", "Score")
score5 = InputBox("Enter Score 5", "Score")


    If score1 > score2 Then
        temp = score1
        score1 = score2
        score2 = temp
    End If

    If score2 > score3 Then
        temp = score2
        score2 = score3
        score3 = temp
    End If

    If score3 > score4 Then
        temp = score3
        score3 = score4
        score4 = temp
    End If

    If score4 > score5 Then
        temp = score4
        score4 = score5
        score5 = temp
    End If
CTR = CTR + 1

sum = score2 + score3 + score4
divetotal = sum * dd
total = total + divetotal
picresults.Print divetotal

End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub Cmdtotal_Click()

picresults.Print "Your score after "; CTR; " dives is"; total




End Sub
