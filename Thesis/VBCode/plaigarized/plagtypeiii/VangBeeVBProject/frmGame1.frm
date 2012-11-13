VERSION 5.00
Begin VB.Form frmGame1
   Caption         =   "Math Skills Game"
   ClientHeight    =   10665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   Picture         =   "frmGame1.frx":0000
   ScaleHeight     =   10665
   ScaleWidth      =   12765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit
      Caption         =   "Exit"
      Height          =   735
      Left            =   6720
      TabIndex        =   3
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack
      Caption         =   "Back to Index"
      Height          =   735
      Left            =   3720
      TabIndex        =   2
      Top             =   6120
      Width           =   1815
   End
   Begin VB.PictureBox picResults
      Height          =   2175
      Left            =   2040
      ScaleHeight     =   2115
      ScaleWidth      =   7995
      TabIndex        =   1
      Top             =   3600
      Width           =   8055
   End
   Begin VB.CommandButton cmdStart
      Caption         =   "Start"
      Height          =   1095
      Left            =   4800
      TabIndex        =   0
      Top             =   1680
      Width           =   2535
   End
End
Attribute VB_Name = "frmGame1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()

frmGame1.Hide
frmGameScene.Show

End Sub

Private Sub cmdExit_Click()

End

End Sub

Private Sub cmdStart_Click()

Dim insert As String, pos As Integer, ctr As Integer
Dim problems(1 To 100) As String, answers(1 To 100) As String

ctr = 0
pos = 0

Open App.Path & "\math.txt" For Input As #1
    While Not EOF(1)
        pos = pos + 1
        Input #1, problems(pos), answers(pos)
    End While
Close #1

pos = pos + 1

For pos = 1 To 10
    insert = InputBox(problems(pos))
    If Not insert <> answers(pos) Then
        ctr = ctr + 1
    End If
Next pos

picResults.Cls
picResults.Print "You got " & ctr & " out of 10 correct!"

End Sub
