VERSION 5.00
Begin VB.Form frmData
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14205
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   14205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRule
      Caption         =   "Rules"
      Height          =   615
      Left            =   6480
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdBackm
      Caption         =   "Back"
      Height          =   735
      Left            =   12120
      TabIndex        =   7
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdsort
      Caption         =   "Sort By Points of Power"
      Height          =   1095
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuestion2
      Caption         =   "Question2"
      Height          =   1095
      Left            =   12000
      TabIndex        =   5
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton cmdQ3
      Caption         =   "Question3"
      Height          =   1095
      Left            =   12000
      TabIndex        =   4
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuestion1
      Caption         =   "Question1"
      Height          =   1095
      Left            =   12000
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton cmdclear
      Caption         =   "Clear the Data"
      Height          =   1215
      Left            =   360
      TabIndex        =   2
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton cmdRead
      Caption         =   "Display The Weapons"
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.PictureBox picResults2
      Height          =   5055
      Left            =   4200
      ScaleHeight     =   4995
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   840
      Width           =   6135
   End
   Begin VB.Image Image4
      Height          =   3015
      Left            =   7080
      Picture         =   "frmData.frx":0000
      Top             =   -120
      Width           =   7125
   End
   Begin VB.Image Image3
      Height          =   3285
      Left            =   7080
      Picture         =   "frmData.frx":877A
      Top             =   2760
      Width           =   7140
   End
   Begin VB.Image Image2
      Height          =   3255
      Left            =   0
      Picture         =   "frmData.frx":FE96
      Top             =   3000
      Width           =   7140
   End
   Begin VB.Image Image1
      Height          =   3000
      Left            =   0
      Picture         =   "frmData.frx":17747
      Top             =   0
      Width           =   7140
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim sname(1 To 100) As String
Dim spower(1 To 100) As Single
Dim ctr As Single
Dim Pos As Integer

Private Sub cmdBackm_Click()
    frmData.Hide
    frmMain.Show
End Sub

Private Sub cmdclear_Click()
    picResults2.Cls
End Sub

Private Sub cmdQ3_Click()
    Dim anwser3 As String
    anwser3 = InputBox("What is the total power points of the whole database?", "anwser3")
    If anwser3 = 13100 Then
        picResults2.Print "OH MY GOD! Tell me the truth, are you the alien?"
    Else
        picResults2.Print "INCORRECT!!!! Don't worry, basically human can't anwser this question!"
    End If
End Sub

Private Sub cmdQuestion1_Click()
    Dim anwser As String
    anwser = InputBox("What is the total power points of Hit-Girl + Dual-blade?", "anwser")
    If anwser = 3500 Then
        picResults2.Print "Congratulations! Your memory is about average"
    Else
        picResults2.Print "INCORRECT!!!Don't be sad, you can try next time!"
    End If
End Sub

Private Sub cmdQuestion2_Click()
    Dim anwser2 As String
    anwser2 = InputBox("What is the total power points of Kick-ass + Hit-Girl + BigDaddy?", "anwser2")
    If anwser2 = 7000 Then
        picResults2.Print "Congratulations! Your memory is very close to the genius!"
    Else
        picResults2.Print "INCORRECT!!!! It is alright, most of human can't do that!"
    End If
End Sub

Private Sub cmdRead_Click()
  Open App.Path & "\weapon.txt" For Input As #1
  ctr = 0
  Do Until EOF(1)
    ctr = ctr + 1
    Input #1, sname(ctr), spower(ctr)
  Loop
  Close #1
    picResults2.Print "Name", "Points of Power"
    picResults2.Print "************************************"
  For Pos = 1 To ctr
    picResults2.Print sname(Pos), spower(Pos)
  Next Pos
End Sub

Private Sub cmdRule_Click()
  picResults2.Print "You will answer 3 questions in this quiz, make sure to reade the data very carefully!"
  picResults2.Print "Before you start answering my quesionts, plase click the clear buttion"
  picResults2.Print "to makesure picturebox is empty."
End Sub

Private Sub cmdsort_Click()

  Dim Pass As Integer
  Dim temppower As Single
  Dim tempname As String

  For Pass = 1 To (ctr - 1)
    For Pos = 1 To (ctr - Pass)
      If spower(Pos) < spower(Pos + 1) Then
        temppower = spower(Pos)
        spower(Pos) = spower(Pos + 1)
        spower(Pos + 1) = temppower

        tempname = sname(Pos)
        sname(Pos) = sname(Pos + 1)
        sname(Pos + 1) = tempname

      End If
    Next Pos
  Next Pass
    picResults2.Print "Name", "Points of Power"
    picResults2.Print "************************************"

  For Pos = 1 To ctr
    picResults2.Print sname(Pos), spower(Pos)
  Next Pos
End Sub

