VERSION 5.00
Begin VB.Form frmTrivia 
   Caption         =   "Trivia"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicResults2 
      BackColor       =   &H00C0FFC0&
      Height          =   2895
      Left            =   4560
      ScaleHeight     =   2835
      ScaleWidth      =   3435
      TabIndex        =   3
      Top             =   3240
      Width           =   3495
   End
   Begin VB.CommandButton cmdTrivia 
      BackColor       =   &H000080FF&
      Caption         =   "Trivia"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0FFC0&
      Height          =   2895
      Left            =   240
      ScaleHeight     =   2835
      ScaleWidth      =   3555
      TabIndex        =   1
      Top             =   3240
      Width           =   3615
   End
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Gabe Hamilton"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label lblScoreCard 
      BackStyle       =   0  'Transparent
      Caption         =   "Trivia Score Card"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   8835
      Left            =   -3840
      Picture         =   "frmTrivia.frx":0000
      Top             =   -2160
      Width           =   13710
   End
   Begin VB.Image Image1 
      Height          =   4935
      Left            =   0
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "frmTrivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdHome_Click()
    'takes user back to the Title page
    frmTrivia.Hide
    frmTitle.Show
End Sub

Private Sub cmdTrivia_Click()
    'Golf Trivia
    'displays the user's answer along with the correct answer
    'points are give to the user lowest score wins
    Dim Quest, Answer, UserAnswer As String
    Dim Total, pos, size, I As Integer
    Total = 0
    pos = 0
    size = 18
    Open App.Path & "\TriviaQuestions.txt" For Input As #1
        picResults.Print "Your Answer"; Tab(15); "Correct Answer"; Tab(35); "Score"
        picResults.Print "_____________________________________________"
        picResults.Print
    For I = 1 To 9
        Input #1, Quest, Answer
        pos = pos + 1
        UserAnswer = InputBox(Quest, "Question " & pos & " of " & size)
        If InStr(Answer, " " & UserAnswer & " ") <> 0 Then
            MsgBox Answer, , "CORRECT!"
            Total = Total - 1
        ElseIf InStr(LCase(Answer), " " & UserAnswer & " ") <> 0 Then
            MsgBox Answer, , "CORRECT!"
            Total = Total - 1
        Else
            MsgBox Answer, , "WRONG!"
            Total = Total + 1
        End If
        picResults.Print UserAnswer; Tab(15); Answer; Tab(38); Total
    Next I
        PicResults2.Print "Your Answer"; Tab(15); "Correct Answer"; Tab(35); "Score"
        PicResults2.Print "_____________________________________________"
        PicResults2.Print
    For I = 10 To 18
        Input #1, Quest, Answer
        pos = pos + 1
        UserAnswer = InputBox(Quest, "Question " & pos & " of " & size)
        If InStr(Answer, " " & UserAnswer & " ") <> 0 Then
            MsgBox Answer, , "CORRECT!"
            Total = Total - 1
        ElseIf InStr(LCase(Answer), " " & UserAnswer & " ") <> 0 Then
            MsgBox Answer, , "CORRECT!"
            Total = Total - 1
        Else
            MsgBox Answer, , "WRONG!"
            Total = Total + 1
        End If
        PicResults2.Print UserAnswer; Tab(15); Answer; Tab(38); Total
    Next I
    Close #1
        PicResults2.Print
        PicResults2.Print "YOUR SCORE IS:"; Tab(38); Total
        PicResults2.Print
            If (Total <= -10) Then
                MsgBox "You're a Pro. ", , "Trivia Score Card"
            ElseIf (Total <= -5) Then
                MsgBox "Not Bad.  You Made the Cut. ", , "Trivia Score Card"
            ElseIf (Total <= 0) Then
                MsgBox "Keep your Head up, Amateur. ", , "Trivia Score Card"
            ElseIf (Total <= 5) Then
                MsgBox "You really could use some practice, the driving range is calling your name. ", , "Trivia Score Card"
            ElseIf (Total <= 18) Then
                MsgBox "You should find a new hobby, golf isn't for you. ", , "Trivia Score Card"
            End If
End Sub

