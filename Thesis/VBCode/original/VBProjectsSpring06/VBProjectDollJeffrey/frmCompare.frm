VERSION 5.00
Begin VB.Form frmCompare 
   BackColor       =   &H00FF0000&
   Caption         =   "Compare to Others!!"
   ClientHeight    =   6945
   ClientLeft      =   8160
   ClientTop       =   855
   ClientWidth     =   6600
   BeginProperty Font 
      Name            =   "Mistral"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   6600
   Begin VB.PictureBox picOutput 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5175
      Left            =   240
      ScaleHeight     =   5115
      ScaleWidth      =   6075
      TabIndex        =   2
      Top             =   1560
      Width           =   6135
   End
   Begin VB.CommandButton cmdWR 
      Caption         =   "Compare to World Record"
      Height          =   1095
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdMIAC 
      Caption         =   "Compare to 2005 MIAC Champion"
      Height          =   1095
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdMIAC_Click()
    Dim pos As Integer
    Dim EV(1 To 10), PTS(1 To 10) As String
    Dim MA(1 To 10), SC(1 To 10) As Integer
    Dim overall, pace As Integer
    'open file and store as input #1
    Open App.Path & "\2005MIAC.txt" For Input As #1
    pos = 0
    'read until end of the file and store each into the set variables
    Do Until EOF(1)
        pos = pos + 1
        Input #1, EV(pos), MA(pos), SC(pos), PTS(pos)
    Loop
    'close the file
    Close #1
    overall = 0
    'compute the overall score
    For pos = 1 To 10
        overall = overall + SC(pos)
    Next pos
    'print the results
    picOutput.Cls
    picOutput.Print , "2005 MIAC Champion"
    picOutput.Print "*****************************************"
    picOutput.Print "Event", "Mark", "Score"
    picOutput.Print "*****************************************"
    For pos = 1 To 10
        picOutput.Print EV(pos), MA(pos), SC(pos), PTS(pos)
    Next pos
    picOutput.Print "*****************************************"
    picOutput.Print "Total", , overall
    picOutput.Print "*****************************************"
    'if they score less that 2005 champ then show how far off they were otherwise show that they have outscored the champ
    If overall > score Then
        pace = overall - score
        picOutput.Print "You are off the 2005 MIAC champ by", pace
    Else
        pace = score - overall
        picOutput.Print "You scored ahead of the 2005 MIAC champ by", pace
                MsgBox "CONGRATULATIONS! YOU OUTSCORED THE 2005 MIAC CHAMPOIN!!!", vbExclamation, "WOW!"
    End If
End Sub
Private Sub cmdWR_Click()
    Dim pos As Integer
    Dim EV(1 To 10), PTS(1 To 10) As String
    Dim MA(1 To 10), SC(1 To 10) As Integer
    Dim overall, pace2 As Integer
    'open file and store as input #1
    Open App.Path & "\WR.txt" For Input As #1
    pos = 0
    'run until end of file storing data into set variables
    Do Until EOF(1)
        pos = pos + 1
        Input #1, EV(pos), MA(pos), SC(pos), PTS(pos)
    Loop
    Close #1
    overall = 0
    'compute the overall score
    For pos = 1 To 10
        overall = overall + SC(pos)
    Next pos
    'print the outcome
    picOutput.Cls
    picOutput.Print , "World Record"
    picOutput.Print "*****************************************"
    picOutput.Print "Event", "Mark", "Score"
    picOutput.Print "*****************************************"
    For pos = 1 To 10
        picOutput.Print EV(pos), MA(pos), SC(pos), PTS(pos)
    Next pos
    picOutput.Print "*****************************************"
    picOutput.Print "Total", , overall
    picOutput.Print "*****************************************"
    picOutput.Print "You are off the World Record by", pace2
    'if world record was missed, display how much it was missed by otherwise display they set the new world record
    If overall > score Then
        pace2 = overall - score
        picOutput.Print "You are off the World Record by", pace2
    Else
        MsgBox "CONGRATULATIONS! YOU HAVE SET THE WORLD RECORD!!!", vbExclamation, "WOW!"
    End If
End Sub
