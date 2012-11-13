VERSION 5.00
Begin VB.Form frmAnswer 
   BackColor       =   &H00404080&
   Caption         =   "Compare your answer to the correct Answer"
   ClientHeight    =   9870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9870
   ScaleWidth      =   12240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNextForm 
      BackColor       =   &H00C0E0FF&
      Caption         =   "More Help"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9495
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdMore 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Keep Studying?"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdCount 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Click if Correct"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6000
      Width           =   1935
   End
   Begin VB.PictureBox picResults4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   9840
      ScaleHeight     =   675
      ScaleWidth      =   435
      TabIndex        =   8
      Top             =   6480
      Width           =   495
   End
   Begin VB.CommandButton cmdImportAnswers 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Import Answers"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdAnswer 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Display Answer"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   8295
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9000
      Width           =   10935
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.PictureBox picResults3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1200
      ScaleHeight     =   1635
      ScaleWidth      =   7635
      TabIndex        =   1
      Top             =   6480
      Width           =   7695
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   1200
      ScaleHeight     =   2235
      ScaleWidth      =   9435
      TabIndex        =   0
      Top             =   2640
      Width           =   9495
   End
   Begin VB.Line Line6 
      X1              =   2160
      X2              =   2400
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line5 
      X1              =   2160
      X2              =   2160
      Y1              =   6360
      Y2              =   6120
   End
   Begin VB.Line Line4 
      X1              =   2520
      X2              =   2160
      Y1              =   5880
      Y2              =   6360
   End
   Begin VB.Line Line3 
      X1              =   2640
      X2              =   3000
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line2 
      X1              =   2640
      X2              =   2880
      Y1              =   2520
      Y2              =   2160
   End
   Begin VB.Line Line1 
      X1              =   4200
      X2              =   2640
      Y1              =   1800
      Y2              =   2520
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   4  'Upward Diagonal
      Height          =   2295
      Left            =   9000
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label lblYour 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Your Answer"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label lblCorrect 
      BackColor       =   &H00C0E0FF&
      Caption         =   "The Correct Answer"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   4  'Upward Diagonal
      Height          =   975
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   4  'Upward Diagonal
      Height          =   855
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   3255
   End
End
Attribute VB_Name = "frmAnswer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'declare the global variables for this form
Dim Answers(1 To 100) As String
Dim AnsNum(1 To 100) As String
Dim Counter As Single
Private Sub cmdAnswer_Click()
Dim Pos As Single
Dim Found As Boolean
Dim CTR As Single
Dim PicAnswer As Single
picResults2.Cls
picResults3.Cls
'this should pull the answer from the first form so the user can compare
picResults3.Print YourAnswer
'this button moves you from the study guide form to the answer form
frmStudyGuide.Visible = False
frmAnswer.Visible = True
'picResults2.Print "You choose question "; PicQuestion; " so chose the same answer number"
'the user can input the number they inputed on the first form to see the correct answer
PicAnswer = InputBox("Enter number " & PicQuestion & " so you get the correct answer for the question you chose")
Found = False
    Do While Found = False And Pos < Counter
    Pos = Pos + 1
        If AnsNum(Pos) = PicQuestion Then
            Found = True
        End If
    Loop
        If Found = True Then
            picResults2.Print AnsNum(Pos), Answers(Pos)
        End If
End Sub
Private Sub cmdBack_Click()
'this button lets you go back to the question or previous form
frmStudyGuide.Show
frmAnswer.Hide
frmWelcome.Hide
frmCompare.Hide
End Sub
Private Sub cmdCount_Click()
'this button allows the user to keep track of how many answers they got right
picResults4.Cls
TrySum = TrySum + 1
picResults4.Print TrySum
End Sub
Private Sub cmdImportAnswers_Click()
'open the file that has the answers in it
Open App.Path & "\Answers.txt" For Input As #2
    Do While Not EOF(2)
        Counter = Counter + 1
        Input #2, AnsNum(Counter), Answers(Counter)
        'picResults2.Print AnsNum(Counter), Answers(Counter)
    Loop
Close #2
picResults3.Cls
'after you import the answers you can choose any button on the form
cmdAnswer.Enabled = True
cmdImportAnswers.Visible = False
End Sub
Private Sub cmdMore_Click()
'this code is for updating the user on the progress they are making as they get questions correct
Select Case TrySum
    Case Is >= 8
    MsgBox "Good Job you are ready for the Test/Quiz"
    Case 4 To 7
    MsgBox "I would practice a few more"
    Case 0 To 3
    MsgBox "You really are not ready yet.  Time for an all-nighter"
End Select
End Sub
Private Sub cmdNextForm_Click()
'this button moves the user to the next form
frmStudyGuide.Hide
frmAnswer.Hide
frmWelcome.Hide
frmCompare.Show
End Sub
Private Sub cmdQuit_Click()
End
End Sub

