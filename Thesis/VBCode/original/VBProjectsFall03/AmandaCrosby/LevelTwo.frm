VERSION 5.00
Begin VB.Form LevelTwo 
   BackColor       =   &H0080C0FF&
   Caption         =   "Level Two"
   ClientHeight    =   4320
   ClientLeft      =   3495
   ClientTop       =   3930
   ClientWidth     =   7485
   LinkTopic       =   "Form2"
   ScaleHeight     =   4320
   ScaleWidth      =   7485
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
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3240
      Width           =   1335
   End
   Begin VB.PictureBox picResultsScore 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   375
      Left            =   5760
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   17
      Top             =   840
      Width           =   855
   End
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
      Height          =   375
      Left            =   240
      MaskColor       =   &H80000007&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3720
      Width           =   1335
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtAnswer 
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   14
      Top             =   3360
      Width           =   4335
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
      Height          =   1095
      Left            =   1800
      ScaleHeight     =   1035
      ScaleWidth      =   2595
      TabIndex        =   10
      Top             =   480
      Width           =   2655
      Begin VB.PictureBox Picture1 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   2775
         TabIndex        =   12
         Top             =   1920
         Width           =   2775
      End
      Begin VB.PictureBox Picture2 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   3735
         TabIndex        =   11
         Top             =   1920
         Width           =   3735
      End
   End
   Begin VB.CommandButton cmdQ4L2 
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
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdQ3L2 
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
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdQ2L2 
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
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdQ1L2 
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
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1335
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
      Height          =   1575
      Left            =   1800
      ScaleHeight     =   1515
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   1560
      Width           =   5295
   End
   Begin VB.CommandButton cmdLevelTwoDone 
      BackColor       =   &H00FF8080&
      Caption         =   "Go Back To Main Menu"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
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
      Left            =   5280
      TabIndex        =   21
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label lblQ3 
      BackColor       =   &H0080C0FF&
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
      Left            =   4560
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblQ4 
      BackColor       =   &H0080C0FF&
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
      Left            =   4560
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   19
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
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
      Left            =   1800
      TabIndex        =   13
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Level Two"
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
      Left            =   2040
      TabIndex        =   9
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblQ1 
      BackColor       =   &H0080C0FF&
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
      Left            =   4560
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblQ2 
      BackColor       =   &H0080C0FF&
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
      Left            =   4560
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "LevelTwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Level Two
'Asks the user questions and gives a score for correct answers
Option Explicit
Dim Question As Integer
Private Sub cmdQuitLevelTwo_Click()
    Form1.Show
    LevelOne.Hide
    LevelTwo.Hide
    LevelThree.Hide
End Sub

Private Sub cmdAnswer_Click()
picResultsText.Cls
Dim Answer As String
If Question = 0 Then
    picResultsText.Cls
    picResultsText.Print "You have already answered this question."
End If
If Question = 1 Then
    Question = 0
    Answer = LCase(txtAnswer)
    If Answer = "four" Then  'if answer is correct,
            LevelTwoScore = LevelTwoScore + 10  'gives points, clears pic, tells correct, thumbs up, clear button
            picResults.Cls
            picResultsText.Print "Correct!"
            picResultsText.Print "Each eighth note is worth half a beat."
            picResultsText.Print "So there are four beats shown here."
            picResults.Picture = LoadPicture(PATH & "Images\ThumbsUp.gif")
            cmdQ1L2.Visible = False
        Else    'if first answer is not correct, asks again
            picResultsText.Print "Sorry.  Incorrect."
            picResultsText.Print "Hint: Each eighth rest is worth half a beat"
            cmdQ1L2.Visible = False
    End If
End If
picResultsScore.Print LevelOneScore
If cmdQ4L2.Visible = False And cmdQ3L2.Visible = False And cmdQ2L2.Visible = False And cmdQ1L2.Visible = False Then
        'goes back to the main form if all the questions have been answered
End If

If Question = 2 Then
    Question = 0        'prevents user from answering twice
    Answer = LCase(txtAnswer) 'gets answer from user
    If Answer = "c" Then  'if answer is correct, gives 10 pts, clears pic, and gets rid of question button
            LevelTwoScore = LevelTwoScore + 10
            picResults.Cls
            picResultsText.Print "Correct!  This is Middle C";
            picResults.Picture = LoadPicture(PATH & "Images\ThumbsUp.gif")
            cmdQ2L2.Visible = False
        Else    'if first answer is not correct, asks again
            picResultsText.Print "Sorry.  Incorrect"
            cmdQ2L2.Visible = False
    End If
End If
picResultsScore.Cls
picResultsScore.Print LevelOneScore
If cmdQ4L2.Visible = False And cmdQ3L2.Visible = False And cmdQ2L2.Visible = False And cmdQ1L2.Visible = False Then
        'goes back to the main form if all the questions have been answered
    cmdLevelTwoDone.Visible = True
End If

If Question = 3 Then
    Question = 0        'prevents user from answering twice
    Answer = LCase(txtAnswer) 'gets answer from user
    If Answer = "four" Then  'if answer is correct, gives 10 pts, clears pic, and gets rid of question button
            LevelTwoScore = LevelTwoScore + 10
            picResults.Cls
            picResultsText.Print "Correct!  Quarter Rests are worth one beat!";
            picResults.Picture = LoadPicture(PATH & "Images\ThumbsUp.gif")
            cmdQ3L2.Visible = False
        Else    'if first answer is not correct, asks again
            picResultsText.Print "Sorry.  Incorrect."
            cmdQ3L2.Visible = False
    End If
End If
picResultsScore.Cls
picResultsScore.Print LevelOneScore
If cmdQ4L2.Visible = False And cmdQ3L2.Visible = False And cmdQ2L2.Visible = False And cmdQ1L2.Visible = False Then
            'goes back to the main form if all the questions have been answered
    cmdLevelTwoDone.Visible = True
End If

If Question = 4 Then
    Question = 0        'prevents user from answering twice
    Answer = LCase(txtAnswer)  'gets answer from user
    If Answer = "whole" Then  'if answer is correct, gives 10 pts, clears pic, and gets rid of question button
            LevelTwoScore = LevelTwoScore + 10
            picResults.Cls
            picResultsText.Print "Correct!  This is whole rest."
            picResults.Picture = LoadPicture(PATH & "Images\ThumbsUp.gif")
            cmdQ4L2.Visible = False
        Else
            picResultsText.Print "Sorry Incorrect."
            cmdQ4L2.Visible = False
    End If
End If
picResultsScore.Cls                     'clears then prints score for level 2
picResultsScore.Print LevelTwoScore
If cmdQ4L2.Visible = False And cmdQ3L2.Visible = False And cmdQ2L2.Visible = False And cmdQ1L2.Visible = False Then
                'goes back to the main form if all the questions have been answered
    cmdLevelTwoDone.Visible = True
End If
End Sub

Private Sub cmdLevelOneDone_Click()
    Select Case LevelTwoScore
        Case Is = 40     'if the use passed the level perfectly
            MsgBox "Congratulations!  You are a music genious!"
            LevelTwoPassed = True
        Case Is = 30
            MsgBox "Great Job!  You passed Level Two"
            LevelTwoPassed = True
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

Private Sub cmdLevelTwoDone_Click()
    Select Case LevelOneScore
        Case Is = 40     'if the use passed the level perfectly
            MsgBox "Congratulations!  You are a music genious!"
            LevelTwoPassed = True
        Case Is = 30
            MsgBox "Great Job!  You passed Level Two"
            LevelTwoPassed = True
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

Private Sub cmdQ1L2_Click()
    picResultsText.Cls
    picResults.Picture = LoadPicture(PATH & "Images\Music\8threst.gif") 'loads 8th rest pic
    picResultsText.Print "How many beats are shown here (in 4/4 time)?" 'tells user what to do
    picResultsText.Print "(Spell out the number)."
    Question = 1        'Tells comp that we are on Question #1
    lblQ1.Visible = True        'displays the Question one Label
    lblQ2.Visible = False
    lblQ3.Visible = False
    lblQ4.Visible = False
End Sub

Private Sub cmdQ2L2_Click()
    picResultsText.Cls
    picResults.Picture = LoadPicture(PATH & "Images\Music\MiddleC.gif") 'loads middle C pic
    picResultsText.Print "What is the name of this note?" 'tells user what to do
    picResultsText.Print "Hint: Middle ___"
    Question = 2        'Tells comp that we are on Question #1
    lblQ2.Visible = True
    lblQ1.Visible = False
    lblQ3.Visible = False
    lblQ4.Visible = False
End Sub

Private Sub cmdQ3L3_Click()
    picResultsText.Cls
    picResults.Picture = LoadPicture(PATH & "Images\Music\quartrest.gif") 'loads quarter rest pic
    picResultsText.Print "How many beats are in this picture (in 4/4 tiem)?" 'tells user what to do
    Question = 3        'Tells comp that we are on Question #1
    lblQ3.Visible = True
    lblQ2.Visible = False
    lblQ1.Visible = False
    lblQ4.Visible = False
End Sub

Private Sub cmdQ4L2_Click()
    picResultsText.Cls
    picResults.Picture = LoadPicture(PATH & "Images\Music\wholerest.gif") 'loads treble clef pic
    picResultsText.Print "This is called a _____ rest."    'tells user what to do
    Question = 4        'Tells comp that we are on Question #1
    lblQ4.Visible = True
    lblQ2.Visible = False
    lblQ3.Visible = False
    lblQ1.Visible = False
End Sub

Private Sub cmdQ3L2_Click()
    picResultsText.Cls
    picResults.Picture = LoadPicture(PATH & "Images\Music\quartrest.gif") 'loads quarter rest pic
    picResultsText.Print "How many beats are in this picture (in 4/4 time)?" 'tells user what to do
    picResultsText.Print "(Spell out your answer)"
    Question = 3        'Tells comp that we are on Question #1
    lblQ3.Visible = True
    lblQ2.Visible = False
    lblQ1.Visible = False
    lblQ4.Visible = False
End Sub

Private Sub Command1_Click()
    Form1.Show
    LevelOne.Hide
    LevelTwo.Hide
End Sub

Private Sub Command2_Click()
    LevelTwoScore = 0
    cmdQ1L2.Visible = True
    cmdQ2L2.Visible = True
    cmdQ3L2.Visible = True
    cmdQ4L2.Visible = True
    picResultsText.Cls
    picResults.Cls
    cmdLevelTwoDone.Visible = False
End Sub
