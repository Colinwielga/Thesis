VERSION 5.00
Begin VB.Form frmFinal 
   BackColor       =   &H00000040&
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   13065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Leave"
      Height          =   615
      Left            =   10680
      TabIndex        =   3
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Timer Timer3 
      Interval        =   2000
      Left            =   720
      Top             =   4080
   End
   Begin VB.Timer Timer4 
      Interval        =   1500
      Left            =   720
      Top             =   2520
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   840
      Top             =   1200
   End
   Begin VB.PictureBox picScore 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   4440
      ScaleHeight     =   435
      ScaleWidth      =   2835
      TabIndex        =   2
      Top             =   4560
      Width           =   2895
   End
   Begin VB.PictureBox picHandicap 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   2880
      Width           =   2415
   End
   Begin VB.PictureBox picUser 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   5160
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   3420
      Left            =   9240
      Picture         =   "frmFinal.frx":0000
      Top             =   840
      Width           =   3315
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Your Statistics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "frmFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Hand As Integer
Dim accuracy As Integer
Dim ta As Integer
'this form calculates the users score or accuracy and displays his/her name, handicaps, and
'score in order seperated by time

Private Sub Command1_Click()
    End
End Sub

Private Sub Timer1_Timer()

    picUser.Print User
    Timer1.Enabled = False
    
End Sub

Private Sub Timer4_Timer()
    If frmOption.ckOne.Value = 1 Then
        Hand = Hand + 1
    End If
    If frmOption.ckMultiple.Value = 1 Then
        Hand = Hand + 1
    End If
    If frmOption.ckShow.Value = 1 Then
        Hand = Hand + 1
    End If
    If frmOption.ckAns.Value = 1 Then
        Hand = Hand + 1
    End If
    picHandicap.Print "Used "; Hand; " handicaps..."
    Timer4.Enabled = False
End Sub

Private Sub Timer3_Timer()
    For Pos = 1 To 81
        If Puz(Pos) = Correct(Pos) Then
            accuracy = accuracy + 1
        End If
    Next Pos
    
    'ta = accuracy / 81
    picScore.Print "Had a total accuracy of "; FormatPercent(accuracy / 81); "."
    Timer3.Enabled = False
    MsgBox accuracy
End Sub
