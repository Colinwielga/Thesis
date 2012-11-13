VERSION 5.00
Begin VB.Form LevelOne 
   BackColor       =   &H008080FF&
   Caption         =   "Level One"
   ClientHeight    =   3615
   ClientLeft      =   3195
   ClientTop       =   3930
   ClientWidth     =   8760
   LinkTopic       =   "Form2"
   ScaleHeight     =   3615
   ScaleWidth      =   8760
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Re-Start Level"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      MaskColor       =   &H80000007&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdAnswer 
      BackColor       =   &H00FF8080&
      Caption         =   "Answer Question"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtAnswer 
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   12
      Top             =   2640
      Width           =   3375
   End
   Begin VB.PictureBox picResultsScore 
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox picResultsText 
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3480
      ScaleHeight     =   1635
      ScaleWidth      =   4875
      TabIndex        =   8
      Top             =   600
      Width           =   4935
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1800
      ScaleHeight     =   1635
      ScaleWidth      =   1635
      TabIndex        =   5
      Top             =   600
      Width           =   1695
      Begin VB.PictureBox Picture2 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   3735
         TabIndex        =   7
         Top             =   1920
         Width           =   3735
      End
      Begin VB.PictureBox Picture1 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   2775
         TabIndex        =   6
         Top             =   1920
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdQ4L1 
      BackColor       =   &H00FF8080&
      Caption         =   "Question Four"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdQ3L1 
      BackColor       =   &H00FF8080&
      Caption         =   "Question Three"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdQ2L1 
      BackColor       =   &H00FF8080&
      Caption         =   "Question Two"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdQ1L1 
      BackColor       =   &H00FF8080&
      Caption         =   "Question One"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Quit this Level"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdLevelOneDone 
      BackColor       =   &H00FF8080&
      Caption         =   "Go back to main menu"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H008080FF&
      Caption         =   "Designed by: Amanda Crosby"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   21
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H008080FF&
      Caption         =   "Type answer here:"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   18
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Level One"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H008080FF&
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblQ1 
      BackColor       =   &H008080FF&
      Caption         =   "Question One"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblQ4 
      BackColor       =   &H008080FF&
      Caption         =   "Question Four"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblQ3 
      BackColor       =   &H008080FF&
      Caption         =   "Question Three"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblQ2 
      BackColor       =   &H008080FF&
      Caption         =   "Question Two"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "LevelOne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Level One
'Asks the user questions and gives scores for correct answers
Option Explicit
Dim Question As Integer


Private Sub cmdAnswer_Click()
Dim Answer As String
If Question = 0 Then
    picResultsText.Cls
    picResultsText.Print "You have already answered this question."
End If
If Question = 1 Then
    Question = 0
    Answer = LCase(txtAnswer)
    If Answer = "treble clef" Then  'if answer is correct,
            LevelOneScore = LevelOneScore + 10  'gives points, clears pic, tells correct, thumbs up, clear button
            picResults.Cls
            picResultsText.Print "Correct!  Treble Clef";
            picResults.Picture = LoadPicture(PATH & "Images\ThumbsUp.gif")
            cmdQ1L1.Visible = False
        Else    'if first answer is not correct, asks again
            picResultsText.Print "Sorry.  Incorrect... it starts with a 'T'"
            cmdQ1L1.Visible = False
    End If
End If
picResultsScore.Print LevelOneScore
If cmdQ4L1.Visible = False And cmdQ3L1.Visible = False And cmdQ2L1.Visible = False And cmdQ1L1.Visible = False Then
        'goes back to the main form if all the questions have been answered
    cmdLevelOneDone.Visible = True
End If

If Question = 2 Then
    Question = 0        'prevents user from answering twice
    Answer = LCase(txtAnswer) 'gets answer from user
    If Answer = "bass clef" Then  'if answer is correct, gives 10 pts, clears pic, and gets rid of question button
            LevelOneScore = LevelOneScore + 10
            picResults.Cls
            picResultsText.Print "Correct!  Bass Clef";
            picResults.Picture = LoadPicture(PATH & "Images\ThumbsUp.gif")
            cmdQ2L1.Visible = False
        Else    'if first answer is not correct, asks again
            picResultsText.Print "Sorry.  Incorrect... it starts with a 'B'"
            cmdQ2L1.Visible = False
    End If
End If
picResultsScore.Cls
picResultsScore.Print LevelOneScore
If cmdQ4L1.Visible = False And cmdQ3L1.Visible = False And cmdQ2L1.Visible = False And cmdQ1L1.Visible = False Then
        'goes back to the main form if all the questions have been answered
    cmdLevelOneDone.Visible = True
End If

If Question = 3 Then
    Question = 0        'prevents user from answering twice
    Answer = LCase(txtAnswer) 'gets answer from user
    If Answer = "repeat" Then  'if answer is correct, gives 10 pts, clears pic, and gets rid of question button
            LevelOneScore = LevelOneScore + 10
            picResults.Cls
            picResultsText.Print "Correct!  This is a repeat sign!";
            picResults.Picture = LoadPicture(PATH & "Images\ThumbsUp.gif")
            cmdQ3L1.Visible = False
        Else    'if first answer is not correct, asks again
            picResultsText.Print "Sorry.  Incorrect... it starts with a 'R'"
            cmdQ3L1.Visible = False
    End If
End If
picResultsScore.Cls
picResultsScore.Print LevelOneScore
If cmdQ4L1.Visible = False And cmdQ3L1.Visible = False And cmdQ2L1.Visible = False And cmdQ1L1.Visible = False Then
            'goes back to the main form if all the questions have been answered
    cmdLevelOneDone.Visible = True
End If

If Question = 4 Then
    Question = 0        'prevents user from answering twice
    Answer = LCase(txtAnswer)  'gets answer from user
    If Answer = "sharp" Then  'if answer is correct, gives 10 pts, clears pic, and gets rid of question button
            LevelOneScore = LevelOneScore + 10
            picResults.Cls
            picResultsText.Print "Correct!  This is a sharp."
            picResultsText.Print "It raises the note one half step!"
            picResults.Picture = LoadPicture(PATH & "Images\ThumbsUp.gif")
            cmdQ4L1.Visible = False
        Else
            picResultsText.Print "Sorry Incorrect...  it starts with a 'S'"
            cmdQ4L1.Visible = False
    End If
End If

picResultsScore.Cls
picResultsScore.Print LevelOneScore
If cmdQ4L1.Visible = False And cmdQ3L1.Visible = False And cmdQ2L1.Visible = False And cmdQ1L1.Visible = False Then
                'goes back to the main form if all the questions have been answered
    cmdLevelOneDone.Visible = True
End If
End Sub


Private Sub cmdLevelOneDone_Click()
    Select Case LevelOneScore
        Case Is = 40     'if the use passed the level perfectly
            MsgBox "Congratulations!  You are a music genious!"
            LevelOnePassed = True
        Case Is = 30
            MsgBox "Great Job!  You passed Level One"
            LevelOnePassed = True
        Case Is = 20
            MsgBox "So Close!  You did not pass this time.  Try again"
        Case Is < 20
            MsgBox "You did not pass.  You may want to study a little"
        Case Else
            MsgBox "Error"
    End Select
    Form1.Show
    LevelOne.Hide
    LevelTwo.Hide
End Sub

Private Sub cmdQ1L1_Click()
    Dim Answer As String
    picResultsText.Cls
    picResults.Picture = LoadPicture(PATH & "Images\Music\Treble2.gif")  'loads treble clef pic
    picResultsText.Print "What is the name of this symbol?" 'tells user what to do
    Question = 1        'Tells comp that we are on Question #1
    lblQ1.Visible = True
    lblQ2.Visible = False
    lblQ3.Visible = False
    lblQ4.Visible = False
End Sub

Private Sub cmdQ2L1_Click()
    Dim Answer As String
    picResultsText.Cls
    picResultsText.Print "What is the name of this symbol?" 'tells user what to do
    picResults.Picture = LoadPicture(PATH & "Images\Music\BassClef.gif") 'loads bass clef pic
    Question = 2        'Tells comp that we are on Question #1
    lblQ2.Visible = True
    lblQ1.Visible = False
    lblQ3.Visible = False
    lblQ4.Visible = False
End Sub

Private Sub cmdQ3L1_Click()
    Dim Answer As String
    'IS THERE A WAY TO CLEAR THE TXTANSWER BOX?
    picResultsText.Cls
    picResultsText.Print "This is called a ______ sign."   'tells user what to do
    picResults.Picture = LoadPicture(PATH & "Images\Music\RepeatSign.gif") 'loads repeat sign pic
    Question = 3        'Tells comp that we are on Question #1
    lblQ3.Visible = True
    lblQ2.Visible = False
    lblQ1.Visible = False
    lblQ4.Visible = False
End Sub

Private Sub cmdQ4L1_Click()
    Dim Answer As String
    'IS THERE A WAY TO CLEAR THE TXTANSWER BOX?
    picResultsText.Cls
    picResultsText.Print "What is the name of this symbol?" 'tells user what to do
    picResults.Picture = LoadPicture(PATH & "Images\Music\Sharp.gif") 'loads sharp sign pic
    Question = 4        'Tells comp that we are on Question #1
    lblQ4.Visible = True
    lblQ2.Visible = False
    lblQ3.Visible = False
    lblQ1.Visible = False
End Sub

Private Sub Command1_Click()
    Form1.Show
    LevelOne.Hide
    LevelTwo.Hide
End Sub

Private Sub Command2_Click()
    LevelOneScore = 0
    cmdLevelOneDone.Visible = False
    cmdQ1L1.Visible = True
    cmdQ2L1.Visible = True
    cmdQ3L1.Visible = True
    cmdQ4L1.Visible = True
    picResultsText.Cls
    picResults.Cls
End Sub

