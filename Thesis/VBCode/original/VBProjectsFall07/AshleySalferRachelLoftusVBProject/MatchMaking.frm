VERSION 5.00
Begin VB.Form frmMatchMaking 
   BackColor       =   &H00FF00FF&
   Caption         =   "MatchMaking"
   ClientHeight    =   10410
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   ScaleHeight     =   10410
   ScaleWidth      =   13065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFC0FF&
      Caption         =   "I'M SCARED! Click to return to the main page."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton cmdViewPicture 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Click to Review Your Selections and Find Your Match!"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Next Question"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdThird 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdFourth 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdSecond 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdBegin 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Press To Begin"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox picWelcome 
      Height          =   1935
      Left            =   3480
      Picture         =   "MatchMaking.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   1200
      Width           =   6015
      Begin VB.PictureBox picName 
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         ScaleHeight     =   555
         ScaleWidth      =   2595
         TabIndex        =   1
         Top             =   960
         Width           =   2655
      End
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3120
      TabIndex        =   4
      Top             =   3600
      Width           =   6975
   End
End
Attribute VB_Name = "frmMatchMaking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'Declares variables for form
    Dim Question(1 To 5) As String, Answer1(1 To 5) As String, Answer2(1 To 5) As String, Answer3(1 To 5) As String, Answer4(1 To 5) As String
    Dim Counter As Integer


Private Sub cmdBegin_Click()
    'Declares variables for subroutine
    Dim Counter1 As Integer, Counter2 As Integer
    
    'Prints users name in picture box
    picName.Print UserName
    
    'Opens file to retrieve questions and answer options
    Open App.Path & "\Questions.txt" For Input As #1
    Counter1 = 0
    Do Until EOF(1)
        Counter1 = Counter1 + 1
        Input #1, Question(Counter1), Answer1(Counter1), Answer2(Counter1), Answer3(Counter1), Answer4(Counter1)
    Loop
    Close #1
    
    'After questions are answered, freezes buttons
    cmdBegin.Visible = False
    cmdNext.Visible = True
    
End Sub



Private Sub cmdFirst_Click()
    
    'Records the users choice for question 1
    cmdFirst.Enabled = False
    cmdSecond.Enabled = False
    cmdThird.Enabled = False
    cmdFourth.Enabled = False
    UserAnswers(Counter) = Answer1(Counter)
    
End Sub

Private Sub cmdFourth_Click()
    
    'Records the users choice for question 4
    cmdFirst.Enabled = False
    cmdSecond.Enabled = False
    cmdThird.Enabled = False
    cmdFourth.Enabled = False
    UserAnswers(Counter) = Answer4(Counter)
    
End Sub

Private Sub cmdNext_Click()
    
    'Runs through all questions displaying answers for each questions
    Counter = Counter + 1
    cmdFirst.Enabled = True
    cmdSecond.Enabled = True
    cmdThird.Enabled = True
    cmdFourth.Enabled = True
        
        
        If Counter > 5 Then
            cmdNext.Visible = False
            cmdFirst.Visible = False
            cmdSecond.Visible = False
            cmdThird.Visible = False
            cmdFourth.Visible = False
        Else
        
            lblDescription.Caption = Question(Counter)
            If Answer1(Counter) <> "?" Then
                cmdFirst.Visible = True
                cmdFirst.Caption = Answer1(Counter)
            Else
                cmdFirst.Visible = False
            End If
            If Answer2(Counter) <> "?" Then
                cmdSecond.Visible = True
                cmdSecond.Caption = Answer2(Counter)
            Else
                cmdSecond.Visible = False
            End If
            If Answer3(Counter) <> "?" Then
                cmdThird.Visible = True
                cmdThird.Caption = Answer3(Counter)
            Else
                cmdThird.Visible = False
            End If
            If Answer4(Counter) <> "?" Then
                cmdFourth.Visible = True
                cmdFourth.Caption = Answer4(Counter)
            Else
                cmdFourth.Visible = False
            End If
        End If
End Sub

Private Sub cmdQuit_Click()
    
    'Goes to next form to display users match
    frmMatchMaking.Hide
    frmMainPage.Show
    
End Sub

Private Sub cmdSecond_Click()
    
    'Records the users choice for question 3
    cmdFirst.Enabled = False
    cmdSecond.Enabled = False
    cmdThird.Enabled = False
    cmdFourth.Enabled = False
    UserAnswers(Counter) = Answer2(Counter)
    
End Sub

Private Sub cmdThird_Click()
    
    'Records the users choice for question 3
    cmdFirst.Enabled = False
    cmdSecond.Enabled = False
    cmdThird.Enabled = False
    cmdFourth.Enabled = False
    UserAnswers(Counter) = Answer3(Counter)
    
End Sub

Private Sub cmdViewPicture_Click()
    
    'Displays user answers and moves to next form
    Dim I As Integer
    Dim AllAnswers As String
    
    For I = 1 To 5
        AllAnswers = AllAnswers & " " & UserAnswers(I)
    Next I
        MsgBox "You have chosen a match with qualities consisting of:" & AllAnswers, , "Your Chosen Characteristics:"

    frmMatchMaking.Hide
    frmCelebrityPicture.Show
    
End Sub

Private Sub Form_Load()

End Sub
