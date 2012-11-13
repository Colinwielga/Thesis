VERSION 5.00
Begin VB.Form Trivia
   Caption         =   "Chinese Zodiac Trivia"
   ClientHeight    =   11580
   ClientLeft      =   6270
   ClientTop       =   2235
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   Picture         =   "Trivia.frx":0000
   ScaleHeight     =   11580
   ScaleWidth      =   8520
   Begin VB.PictureBox picScore
      BeginProperty Font
         Name            =   "Ravie"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   12
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdReturn
      Caption         =   "I'm bored, get me out of here!"
      BeginProperty Font
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   6720
      TabIndex        =   10
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdNext
      Caption         =   "Next!"
      BeginProperty Font
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6720
      TabIndex        =   9
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtNumber
      BeginProperty Font
         Name            =   "Ravie"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   8
      Top             =   9960
      Width           =   2175
   End
   Begin VB.CommandButton cmdDone
      Caption         =   "Hit me after typing in your answer!"
      BeginProperty Font
         Name            =   "Rockwell Extra Bold"
         Size            =   9
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   5
      Top             =   10800
      Width           =   2655
   End
   Begin VB.TextBox txtName
      BeginProperty Font
         Name            =   "Snap ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   3
      Top             =   9960
      Width           =   2175
   End
   Begin VB.CommandButton cmdStart
      Caption         =   "Start!"
      BeginProperty Font
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6720
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.PictureBox picQuestion
      AutoSize        =   -1  'True
      Height          =   5415
      Left            =   1680
      ScaleHeight     =   5355
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Label LbScore
      Caption         =   "Score:"
      BeginProperty Font
         Name            =   "Ravie"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label3
      Caption         =   "Number:"
      BeginProperty Font
         Name            =   "Ravie"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   7
      Top             =   9120
      Width           =   1575
   End
   Begin VB.Label Label2
      Caption         =   "Name:"
      BeginProperty Font
         Name            =   "Ravie"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Label Lbinstructions
      BackColor       =   &H8000000E&
      Caption         =   "Can you recognize the Zodiac in the picture above? If you can, type in its name and  number!"
      BeginProperty Font
         Name            =   "Ravie"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   4
      Top             =   8160
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label Label1
      BackColor       =   &H00FFFFFF&
      Caption         =   $"Trivia.frx":A0F5
      BeginProperty Font
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   1695
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "Trivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QuestionNum As Integer

Private Sub cmdDone_Click()
Dim AnswerName As String, AnswerNum As Integer
AnswerName = ""
AnswerName = txtName.Text                      'get the input from the user
AnswerNum = txtNumber.Text
If UCase(AnswerName) = zodiac(QuestionNum) And AnswerNum = QuestionNum Then       'check if the answers are both right. if not, present the right answers,
    MsgBox "Good job! That's exactly what I expected!", , "!!!!"
    score = score
    score = score + 1                             'one correct set of questions will gain one score.
    picScore.Cls
    picScore.Print score
Else
    MsgBox "Oops, your answers seem incorrect. The right name of this Zodiac is " & zodiac(QuestionNum) & " and its number is " & QuestionNum, , "Sorry"
    picScore.Cls
    picScore.Print score
End If
cmdDone.Enabled = False
End Sub

Private Sub cmdNext_Click()
Randomize
QuestionNum = Int(12 * Rnd) + 1                    'come up with a random picture
picQuestion.Picture = LoadPicture(App.Path & "\images\" & Names(QuestionNum))
cmdDone.Enabled = True
End Sub

Private Sub cmdReturn_Click()
Home.Visible = True
Trivia.Visible = False
If score > 3 Then                                  'bonus for score more than one
    Home.Pic = LoadPicture(App.Path & "\images\shengxiao.jpg")
    MsgBox "You did a good job!", , "Good for you!"
ElseIf score = 0 Then
    MsgBox "Do you want to check with the picture on the home form, and come back again?", , "Better luck next time!"
End If
cmdStart.Visible = True
cmdStart.Visible = True
End Sub

Private Sub cmdStart_Click()
cmdNext.Visible = True                      'initializing the score
If score <> 0 Then                          'warns the user that the score has been initialized
    MsgBox "All previous scores have been initiallized", , "Warning"
End If
score = 0
Randomize                                    'come up with a random picture for user to recognize.
QuestionNum = Int(12 * Rnd) + 1
picQuestion.Picture = LoadPicture(App.Path & "\images\" & Names(QuestionNum))
Lbinstructions.Visible = True                'this label shows instructions for user to answer the questions, and is show after a picture is printed.
cmdStart.Visible = False
cmdNext.Visible = True
cmdDone.Enabled = True
End Sub

