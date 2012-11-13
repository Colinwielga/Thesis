VERSION 5.00
Begin VB.Form frmPlay 
   BackColor       =   &H00000000&
   Caption         =   "PacMan"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   394
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   733
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000080FF&
      Cancel          =   -1  'True
      Caption         =   "Quit :["
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdGet 
      BackColor       =   &H00FF80FF&
      Caption         =   "Get 1st Question"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00FF80FF&
      Caption         =   "Start!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H0000FFFF&
      Caption         =   "Enter :]"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtAnswer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picQuestions 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   840
      ScaleHeight     =   3315
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.Label lblEnter 
      BackColor       =   &H00000000&
      Caption         =   "Enter Letter Here -->"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   4800
      Visible         =   0   'False
      Width           =   2775
   End
End
Attribute VB_Name = "frmPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim triviaQs(1 To 100) As String, answerA(1 To 100) As String, answerB(1 To 100) As String, answerC(1 To 100) As String, answerD(1 To 100) As String, correct(1 To 100) As String, correctAnswer(1 To 100) As String, InputAnswer As String, CTR As Integer, Pos As Integer, Wrong As Boolean
Private Sub cmdEnter_Click()

    InputAnswer = txtAnswer.Text
    Pos = Pos + 1
    If LCase(correct(Pos)) = LCase(InputAnswer) Then
        txtAnswer.Text = ""
        picQuestions.Cls
        picQuestions.Print "Correct!"
        picQuestions.Print "     "
        picQuestions.Print triviaQs(Pos + 1)
        picQuestions.Print "     "
        picQuestions.Print "A) " & answerA(Pos + 1)
        picQuestions.Print "B) " & answerB(Pos + 1)
        picQuestions.Print "C) " & answerC(Pos + 1)
        picQuestions.Print "D) " & answerD(Pos + 1)
            If Pos >= CTR Then
                picQuestions.Cls
                picQuestions.Print "Correct!"
                picQuestions.Print "You got " & CTR & " out of " & CTR & " correct!"
            End If
    Else
        picQuestions.Cls
        picQuestions.Print "Incorrect! The correct answer is " & correctAnswer(Pos)
        picQuestions.Print "You got " & Pos - 1 & " out of " & CTR & " correct."
    End If

    
End Sub


Private Sub cmdGet_Click()

    picQuestions.Print triviaQs(1)
    picQuestions.Print "A) " & answerA(1)
    picQuestions.Print "B) " & answerB(1)
    picQuestions.Print "C) " & answerC(1)
    picQuestions.Print "D) " & answerD(1)
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdStart_Click()
    


    Open App.Path & "\Trivia.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, triviaQs(CTR), answerA(CTR), answerB(CTR), answerC(CTR), answerD(CTR), correct(CTR), correctAnswer(CTR)
    Loop
    Close #1


    
    cmdStart.Visible = False
    picQuestions.Visible = True
    cmdGet.Visible = True
    lblEnter.Visible = True
    txtAnswer.Visible = True
    cmdEnter.Visible = True
    cmdBack.Visible = True
    cmdQuit.Visible = True
End Sub

Private Sub List1_Click()

End Sub
