VERSION 5.00
Begin VB.Form frmVerbTest 
   BackColor       =   &H00000080&
   Caption         =   "Lingua Vivens - Student Tests - Verb Forms"
   ClientHeight    =   9585
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19320
   LinkTopic       =   "Form1"
   ScaleHeight     =   9585
   ScaleWidth      =   19320
   Begin VB.PictureBox picPrinciple 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   2955
      TabIndex        =   41
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton cmdLogOut 
      BackColor       =   &H00808080&
      Caption         =   "Log Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16200
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   8160
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00808080&
      Caption         =   "Return to Student Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   8160
      Width           =   2175
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00000080&
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   8160
      Width           =   2175
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00000080&
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   8160
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H00000080&
      Caption         =   "Submit and Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8160
      Width           =   2175
   End
   Begin VB.Frame fraMood 
      BackColor       =   &H00000080&
      Caption         =   "Select a Mood"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   600
      TabIndex        =   23
      Top             =   2400
      Width           =   2415
      Begin VB.OptionButton optInfinitive 
         BackColor       =   &H00000080&
         Caption         =   "Infinitive"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   240
         TabIndex        =   31
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton optParticiple 
         BackColor       =   &H00000080&
         Caption         =   "Participle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   240
         TabIndex        =   30
         Top             =   1680
         Width           =   1575
      End
      Begin VB.OptionButton optImperative 
         BackColor       =   &H00000080&
         Caption         =   "Imperative"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   240
         TabIndex        =   29
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton optSubjungtive 
         BackColor       =   &H00000080&
         Caption         =   "Subjungtive"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optIndicative 
         BackColor       =   &H00000080&
         Caption         =   "Indicative"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.PictureBox picVerbsTested 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   8160
      ScaleHeight     =   7515
      ScaleWidth      =   10155
      TabIndex        =   9
      Top             =   360
      Width           =   10215
   End
   Begin VB.Frame fraVoice 
      BackColor       =   &H00000080&
      Caption         =   "Select a voice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   5640
      TabIndex        =   8
      Top             =   2400
      Width           =   2415
      Begin VB.OptionButton optPassive 
         BackColor       =   &H00000080&
         Caption         =   "Passive"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optActive 
         BackColor       =   &H00000080&
         Caption         =   "Active"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fraPerson 
      BackColor       =   &H00000080&
      Caption         =   "Select a Person"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   600
      TabIndex        =   7
      Top             =   5760
      Width           =   2415
      Begin VB.OptionButton optFirstP 
         BackColor       =   &H00000080&
         Caption         =   "First Person"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   240
         TabIndex        =   39
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optNoPerson 
         BackColor       =   &H00000080&
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   240
         TabIndex        =   33
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optThirdP 
         BackColor       =   &H00000080&
         Caption         =   "Third Person"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optSecondP 
         BackColor       =   &H00000080&
         Caption         =   "Second Person"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   1750
      End
   End
   Begin VB.Frame fraConjugation 
      BackColor       =   &H00000080&
      Caption         =   "Select a Conjugation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   5640
      TabIndex        =   6
      Top             =   5760
      Width           =   2415
      Begin VB.OptionButton optFirst 
         BackColor       =   &H00000080&
         Caption         =   "First Conjugation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   2000
      End
      Begin VB.OptionButton optFourth 
         BackColor       =   &H00000080&
         Caption         =   "Fourth Conjugation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   2200
      End
      Begin VB.OptionButton optThirdIO 
         BackColor       =   &H00000080&
         Caption         =   "Third -IO Conjugation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   2205
      End
      Begin VB.OptionButton optThird 
         BackColor       =   &H00000080&
         Caption         =   "Third Conjugation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   2085
      End
      Begin VB.OptionButton optSecond 
         BackColor       =   &H00000080&
         Caption         =   "Second Conjugation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   2085
      End
   End
   Begin VB.Frame fraNumber 
      BackColor       =   &H00000080&
      Caption         =   "Select a Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3120
      TabIndex        =   5
      Top             =   5760
      Width           =   2415
      Begin VB.OptionButton optNoNumber 
         BackColor       =   &H00000080&
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   240
         TabIndex        =   32
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optPlural 
         BackColor       =   &H00000080&
         Caption         =   "Plural"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton optSingular 
         BackColor       =   &H00000080&
         Caption         =   "Singular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame fraTense 
      BackColor       =   &H00000080&
      Caption         =   "Select a Tense"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   3120
      TabIndex        =   4
      Top             =   2400
      Width           =   2415
      Begin VB.OptionButton optFuturePerfect 
         BackColor       =   &H00000080&
         Caption         =   "Future Perfect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   240
         TabIndex        =   15
         Top             =   2640
         Width           =   1575
      End
      Begin VB.OptionButton optPluPerfect 
         BackColor       =   &H00000080&
         Caption         =   "PluPerfect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton optPerfect 
         BackColor       =   &H00000080&
         Caption         =   "Perfect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton optFuture 
         BackColor       =   &H00000080&
         Caption         =   "Future"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton optImperfect 
         BackColor       =   &H00000080&
         Caption         =   "Imperfect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton optPresent 
         BackColor       =   &H00000080&
         Caption         =   "Present"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.PictureBox PicGrade 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.PictureBox picVerb 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Verb Test"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   36
         Charset         =   255
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   9120
      TabIndex        =   43
      Top             =   8040
      Width           =   3135
   End
   Begin VB.Label lblPrinciplesParts 
      BackStyle       =   0  'Transparent
      Caption         =   "Principle Parts:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4920
      TabIndex        =   42
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label lblGrade 
      BackStyle       =   0  'Transparent
      Caption         =   "# Correct/ # Tested"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblVerb 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmVerbTest.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   9015
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   18855
   End
End
Attribute VB_Name = "frmVerbTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variables for collecting data from the user
Dim answerTense As Integer, ATenseName As String
Dim answerMood As Integer, AMoodName As String
Dim answerNumber As Integer, ANumName As String
Dim answerPerson As Integer, APersonName As String
Dim answerConjugation As Integer, AConjName As String
Dim answerVoice As Integer, AVoiceName As String
'Generated test variable
Dim testMood As Integer, tMoodName As String
Dim testTense As Integer, tTenseName As String
Dim testNumber As Integer, tNumName As String
Dim testPerson As Integer, tPersonName As String
Dim testConjugation As Integer, tConjName As String
Dim testVoice As Integer, tVoiceName As String
Dim VerbType As Integer
Dim rndVerb As String
Dim TestVerb As String
Dim previousVerb As String
'Variables for collecting grading information
Dim maxGrade As Integer
Dim NumCorrect As Integer
Dim NumWrong As Integer
'A test verabile to determine if the user has not answered on of the field and to exit the outermost subroutine
Dim ExitSUB As Boolean

Public Sub GetAnswerData()
    'Gets the data for from the user with respect to his answer
    'Gets the data for Tense and rests the option buttons
    If optPresent.Value = True Then
        answerTense = 1
        ATenseName = "Pres."
        optPresent.Value = False
    ElseIf optImperfect.Value = True Then
        answerTense = 2
        ATenseName = "Imperf."
        optImperfect.Value = False
    ElseIf optFuture.Value = True Then
        answerTense = 3
        ATenseName = "Fut."
        optFuture.Value = False
    ElseIf optPerfect.Value = True Then
        answerTense = 4
        ATenseName = "Perf."
        optPerfect.Value = False
    ElseIf optPluPerfect.Value = True Then
        answerTense = 5
        ATenseName = "PluPerf."
        optPluPerfect.Value = False
    ElseIf optFuturePerfect.Value = True Then
        answerTense = 6
        ATenseName = "FutPerf."
        optFuturePerfect.Value = False
    Else
        MsgBox "Please select a tense before continuing"
        ExitSUB = True
        Exit Sub
    End If
    'Mood Data
    If optIndicative.Value = True Then
        answerMood = 1
        AMoodName = "Indic."
        optIndicative.Value = False
    ElseIf optSubjungtive.Value = True Then
        answerMood = 2
        AMoodName = "Subj."
        optSubjungtive.Value = False
    ElseIf optImperative.Value = True Then
        answerMood = 3
        AMoodName = "Imper."
        optImperative.Value = False
    ElseIf optParticiple.Value = True Then
        answerMood = 4
        AMoodName = "Part."
        optParticiple.Value = False
    ElseIf optInfinitive.Value = True Then
        answerMood = 5
        AMoodName = "Infin."
        optInfinitive.Value = False
    Else
        MsgBox "Please select a mood before continuing."
        ExitSUB = True
        Exit Sub
    End If
    'Number Data
    If optSingular.Value = True Then
        answerNumber = 1
        ANumName = "S."
        optSingular.Value = False
    ElseIf optPlural.Value = True Then
        answerNumber = 2
        ANumName = "P."
        optPlural.Value = False
    ElseIf optNoNumber.Value = True Then
        answerNumber = 3
        ANumName = ""
        optNoNumber.Value = False
    Else
        MsgBox "Please select a number before continuing"
        ExitSUB = True
        Exit Sub
    End If
    'Person Data
    If optFirstP.Value = True Then
        answerPerson = 1
        APersonName = "1st,"
        optFirstP.Value = False
    ElseIf optSecondP.Value = True Then
        answerPerson = 2
        APersonName = "2nd,"
        optSecondP.Value = False
    ElseIf optThirdP.Value = True Then
        answerPerson = 3
        APersonName = "3rd,"
        optThirdP.Value = False
    ElseIf optNoPerson.Value = True Then
        answerPerson = 4
        APersonName = ""
        optNoPerson.Value = False
    Else
        MsgBox "Please select a person before continuing"
        ExitSUB = True
        Exit Sub
    End If
    'Conjugation Data
    If optFirst.Value = True Then
        answerConjugation = 1
        AConjName = "1st"
        optFirst.Value = False
    ElseIf optSecond.Value = True Then
        answerConjugation = 2
        AConjName = "2nd"
        optSecond.Value = False
    ElseIf optThird.Value = True Then
        answerConjugation = 3
        AConjName = "3rd"
        optThird.Value = False
    ElseIf optThirdIO.Value = True Then
        answerConjugation = 4
        AConjName = "3rd -IO"
        optThirdIO.Value = False
    ElseIf optFourth.Value = True Then
        answerConjugation = 5
        AConjName = "4th"
        optFourth.Value = False
    Else
        MsgBox "Please select a conjugation before continuing"
        ExitSUB = True
        Exit Sub
    End If
    'Voice Data
    If optActive.Value = True Then
        answerVoice = 1
        AVoiceName = "Act."
        optActive.Value = False
    ElseIf optPassive.Value = True Then
        answerVoice = 2
        AVoiceName = "Pass."
        optPassive.Value = False
    Else
        MsgBox "Please select a voice before continuing"
        ExitSUB = True
        Exit Sub
    End If
    
End Sub

Public Sub IndicativeSequence()
        'Randomize viual basic random number generator
        Randomize
        'Gives test tense a value (this will also determine the voice of the verb)
        testTense = Int((14 - 1 + 1) * Rnd + 1)
        ' deals with the duplication of the future tense across two conjugation
        Select Case testConjugation
            Case 1, 2 ' If first or seond
                If testTense = 7 Then ' then if 3/4 conj. future
                    testTense = 5 ' switch to 1/2 future conjugation of the same voice
                ElseIf testTense = 8 Then
                    testTense = 6
                End If
            Case 3 To 5 ' If third or fourth
                If testTense = 5 Then ' if 1/2 future conj.
                    testTense = 7 'switch to 3/4 fut conj of same voice
                ElseIf testTense = 6 Then
                    testTense = 8
                End If
        End Select
        
        'Deals with Deponent, Semi Deponent, and Defective Verbs
    Select Case testTense
        Case 1, 3, 5, 7, 9, 11, 13 'If active
            If VerbType = 2 Then 'If deponent
                testTense = testTense + 1 'make passive
                testVoice = 2
            ElseIf VerbType = 3 Then 'If semiDeponent
                If testTense = 9 Or testTense = 11 Or testTense = 13 Then ' if perfect sequence
                    testTense = testTense + 1 'make passive
                    testVoice = 2
                End If
            End If
        Case 2, 4, 6, 8, 10, 12, 14 'If Passive
            If VerbType = 4 Then 'If Defective
                testTense = testTense - 1 'make active
            End If
    End Select
        'Main descision and concatination of verb stems with ending
        Select Case testTense
        
            Case 1, 2 'Present active/passive
                If testTense = 1 Then
                    testVoice = 1
                Else
                    testVoice = 2
                End If
                
                If testNumber = 1 Then ' If singular
                    If testPerson = 1 Then ' If first person
                        TestVerb = VerbPresStem(rndVerb) & IFirstS(testTense)
                    ElseIf testPerson = 2 Then 'If second person
                        If testConjugation = 4 And testTense = 2 Then 'If special case of the 3rd-IO conjugation (present passive secon singulat
                            TestVerb = VerbPresStem(rndVerb) & "e" & ISecondS(testTense)
                        Else
                            TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & ISecondS(testTense)
                        End If
                    ElseIf testPerson = 3 Then
                        TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & IThirdS(testTense)
                    End If
                Else
                    If testPerson = 1 Then
                        TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & IFirstP(testTense)
                    ElseIf testPerson = 2 Then
                        TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & ISecondP(testTense)
                    ElseIf testPerson = 3 Then
                        TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & IThirdP(testTense)
                    End If
                End If
            Case 3, 4  'Imperfect active/Passive
                If testTense = 3 Then
                    testVoice = 1
                Else
                    testVoice = 2
                End If
                
               If testNumber = 1 Then
                    If testPerson = 1 Then
                        TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & IFirstS(testTense)
                    ElseIf testPerson = 2 Then
                        TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & ISecondS(testTense)
                    ElseIf testPerson = 3 Then
                        TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & IThirdS(testTense)
                    End If
                Else
                    If testPerson = 1 Then
                        TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & IFirstP(testTense)
                    ElseIf testPerson = 2 Then
                        TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & ISecondP(testTense)
                    ElseIf testPerson = 3 Then
                        TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & IThirdP(testTense)
                    End If
                End If
            Case 5 To 8 'Future active/Passive
                If testTense = 5 Or testTense = 7 Then
                    testVoice = 1
                Else
                    testVoice = 2
                End If
                
                Select Case testConjugation
                    Case 1, 2, 3, 5 'deals with special case of 3rd-IO future verb forms
                        If testNumber = 1 Then
                            If testPerson = 1 Then
                                TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & IFirstS(testTense)
                            ElseIf testPerson = 2 Then
                                TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & ISecondS(testTense)
                            ElseIf testPerson = 3 Then
                                TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & IThirdS(testTense)
                            End If
                        Else
                            If testPerson = 1 Then
                                TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & IFirstP(testTense)
                            ElseIf testPerson = 2 Then
                                TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & ISecondP(testTense)
                            ElseIf testPerson = 3 Then
                                TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & IThirdP(testTense)
                            End If
                        End If
                    Case 4
                        If testNumber = 1 Then
                            If testPerson = 1 Then
                                TestVerb = VerbPresStem(rndVerb) & IFirstS(testTense)
                            ElseIf testPerson = 2 Then
                                TestVerb = VerbPresStem(rndVerb) & ISecondS(testTense)
                            ElseIf testPerson = 3 Then
                                TestVerb = VerbPresStem(rndVerb) & IThirdS(testTense)
                            End If
                        Else
                            If testPerson = 1 Then
                                TestVerb = VerbPresStem(rndVerb) & IFirstP(testTense)
                            ElseIf testPerson = 2 Then
                                TestVerb = VerbPresStem(rndVerb) & ISecondP(testTense)
                            ElseIf testPerson = 3 Then
                                TestVerb = VerbPresStem(rndVerb) & IThirdP(testTense)
                            End If
                        End If
                End Select
            Case 9, 11, 13 'If in the perfect  active sequence
                testVoice = 1
                If testNumber = 1 Then
                    If testPerson = 1 Then
                        TestVerb = VerbPerfStem(rndVerb) & IFirstS(testTense)
                    ElseIf testPerson = 2 Then
                        TestVerb = VerbPerfStem(rndVerb) & ISecondS(testTense)
                    ElseIf testPerson = 3 Then
                        TestVerb = VerbPerfStem(rndVerb) & IThirdS(testTense)
                    End If
                Else
                    If testPerson = 1 Then
                        TestVerb = VerbPerfStem(rndVerb) & IFirstP(testTense)
                    ElseIf testPerson = 2 Then
                        TestVerb = VerbPerfStem(rndVerb) & ISecondP(testTense)
                    ElseIf testPerson = 3 Then
                        TestVerb = VerbPerfStem(rndVerb) & IThirdP(testTense)
                    End If
                End If
            Case 10, 12, 14 'If in the perfect passive sequence
                testVoice = 2
                If testNumber = 1 Then
                    If testPerson = 1 Then
                        TestVerb = VerbPartStem(rndVerb) & IFirstS(testTense)
                    ElseIf testPerson = 2 Then
                        TestVerb = VerbPartStem(rndVerb) & ISecondS(testTense)
                    ElseIf testPerson = 3 Then
                        TestVerb = VerbPartStem(rndVerb) & IThirdS(testTense)
                    End If
                Else
                    If testPerson = 1 Then
                        TestVerb = VerbPartStem(rndVerb) & IFirstP(testTense)
                    ElseIf testPerson = 2 Then
                        TestVerb = VerbPartStem(rndVerb) & ISecondP(testTense)
                    ElseIf testPerson = 3 Then
                        TestVerb = VerbPartStem(rndVerb) & IThirdP(testTense)
                    End If
                End If
                       
        End Select
        'Gives a testable tense to testTense (which is concurrent with the six tense order pres,imprf,fut,perf,pluperf,futperf numbered accordingly)
        Select Case testTense
            Case 1, 2 ' If pres
                testTense = 1 'give testTense pres value
            Case 3, 4 ' if imperf
                testTense = 2 ' give testTense imperf value
            Case 5 To 8
                testTense = 3
            Case 9, 10
                testTense = 4
            Case 11, 12
                testTense = 5
            Case 13, 14
                testTense = 6
        End Select
        
        'Select Case testTense
            'Case 1, 3, 5, 7, 9, 11, 13 ' If active
                'testVoice = 1 ' set testVoce active
            'Case 2, 4, 6, 8, 10, 12, 14 ' If passive
                'testVoice = 2 ' ste testVoice passive
        'End Select
        
    'If Right(IVerbFormName(testTense), 6) = "Active" Then
        'testTense = 1
    'ElseIf Right(IVerbFormName(testTense), 7) = "Passive" Then
        'testTense = 2
    'End If
    
End Sub

Public Sub SubjungtiveSequence()
    'Randomizes the visualbasic randomNumber generator
    Randomize
    'Gives random tense with a subjungtive array range (ie. 8 instead of 14)
    testTense = Int((8 - 1 + 1) * Rnd + 1)
    'Deals with Deponent, Semi Deponent, and Defective Verbs
    Select Case testTense
        Case 1, 3, 5, 7 'If active
            If VerbType = 2 Then 'If deponent
                testTense = testTense + 1
                testVoice = 2
            ElseIf VerbType = 3 Then 'If semiDeponent
                If testTense = 5 Or testTense = 7 Then 'If Completed Action
                    testTense = testTense + 1 'Make passive
                    testVoice = 2
                End If
            End If
        Case 2, 4, 6, 8 'If Passive
            If VerbType = 4 Then 'If Defective
                testTense = testTense - 1 ' make active
            End If
    End Select
            
            
    'Main descision for the subjungtive test verb of tense(testTense)
    Select Case testTense
        Case 1, 2 ' present subjungtive active/passive
            If testTense = 1 Then
                testVoice = 1
            Else
                testVoice = 2
            End If
            
            If testNumber = 1 Then
                If testPerson = 1 Then
                    TestVerb = VerbPresStem(rndVerb) & VowelSubjungtive(testConjugation) & SFirstS(testTense)
                ElseIf testPerson = 2 Then
                    TestVerb = VerbPresStem(rndVerb) & VowelSubjungtive(testConjugation) & SSecondS(testTense)
                ElseIf testPerson = 3 Then
                    TestVerb = VerbPresStem(rndVerb) & VowelSubjungtive(testConjugation) & SThirdS(testTense)
                End If
            Else
                If testPerson = 1 Then
                    TestVerb = VerbPresStem(rndVerb) & VowelSubjungtive(testConjugation) & SFirstP(testTense)
                ElseIf testPerson = 2 Then
                    TestVerb = VerbPresStem(rndVerb) & VowelSubjungtive(testConjugation) & SSecondP(testTense)
                ElseIf testPerson = 3 Then
                    TestVerb = VerbPresStem(rndVerb) & VowelSubjungtive(testConjugation) & SThirdP(testTense)
                End If
            End If
        Case 3, 4 'imperfectSubjungtive
            If testTense = 3 Then
                testVoice = 1
            Else
                testVoice = 2
            End If
            
            If testNumber = 1 Then
                If testPerson = 1 Then
                    TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & SFirstS(testTense)
                ElseIf testPerson = 2 Then
                    TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & SSecondS(testTense)
                ElseIf testPerson = 3 Then
                    TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & SThirdS(testTense)
                End If
            Else
                If testPerson = 1 Then
                    TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & SFirstP(testTense)
                ElseIf testPerson = 2 Then
                    TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & SSecondP(testTense)
                ElseIf testPerson = 3 Then
                    TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & SThirdP(testTense)
                End If
            End If
        Case 5, 7 'perfect sequence subjungtive (active)
            testVoice = 1
            If testNumber = 1 Then
                If testPerson = 1 Then
                    TestVerb = VerbPerfStem(rndVerb) & SFirstS(testTense)
                ElseIf testPerson = 2 Then
                    TestVerb = VerbPerfStem(rndVerb) & SSecondS(testTense)
                ElseIf testPerson = 3 Then
                    TestVerb = VerbPerfStem(rndVerb) & SThirdS(testTense)
                End If
            Else
                If testPerson = 1 Then
                    TestVerb = VerbPerfStem(rndVerb) & SFirstP(testTense)
                ElseIf testPerson = 2 Then
                    TestVerb = VerbPerfStem(rndVerb) & SSecondP(testTense)
                ElseIf testPerson = 3 Then
                    TestVerb = VerbPerfStem(rndVerb) & SThirdP(testTense)
                End If
            End If
        Case 6, 8 'perfect sequence subjungtive passive
            testVoice = 2
            If testNumber = 1 Then
                If testPerson = 1 Then
                    TestVerb = VerbPartStem(rndVerb) & SFirstS(testTense)
                ElseIf testPerson = 2 Then
                    TestVerb = VerbPartStem(rndVerb) & SSecondS(testTense)
                ElseIf testPerson = 3 Then
                    TestVerb = VerbPartStem(rndVerb) & SThirdS(testTense)
                End If
            Else
                If testPerson = 1 Then
                    TestVerb = VerbPartStem(rndVerb) & SFirstP(testTense)
                ElseIf testPerson = 2 Then
                    TestVerb = VerbPartStem(rndVerb) & SSecondP(testTense)
                ElseIf testPerson = 3 Then
                    TestVerb = VerbPartStem(rndVerb) & SThirdP(testTense)
                End If
            End If
    End Select
    
    Select Case testTense 'Gives the proper universal tense to the testTense varialbe
        Case 1, 2 ' if present
            testTense = 1 ' mkae present
        Case 3, 4 ' if imperfect
            testTense = 2 'make imperfect
        Case 5, 6 ' if perfect
            testTense = 4 'make perfect
        Case 7, 8 ' if pluperf
            testTense = 5 ' make pluperf
    End Select
    
    'If Right(SVerbFormName(testTense), 6) = "Active" Then
        'testTense = 1
    'ElseIf Right(SVerbFormName(testTense), 7) = "Passive" Then
        'testTense = 2
    'End If
    
    
End Sub
'Gives imperative ending to the verbs
Public Sub Imperatives()
    'Randomizes visual basics random number generator
    Randomize
    'Sets the appropriate tense for the imperative mood
    testTense = 1 ' present
    testPerson = 2 ' second person
    testVoice = Int((2 - 1 + 1) * Rnd + 1) ' active
    'main descision branch
    If testVoice = 1 Then ' Tests active/passive : if active
        If testNumber = 1 Then ' if singular
            If testConjugation = 1 Or testConjugation = 2 Or testConjugation = 3 Or testConjugation = 5 Then 'If not third-IO
                TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation)
            Else ' Special case for third IO verbs
                TestVerb = VerbPresStem(rndVerb) & "e"
            End If
        Else ' If plural
            TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & "te"
        End If
    ElseIf testVoice = 2 Then 'If passive
        If testNumber = 1 Then
            TestVerb = VerbInfinitive(rndVerb)
            optImperative.Value = True
        Else
            TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & "mini"
        End If
    End If
    
    
End Sub

Public Sub Participles()
    'creates the participle verb forms
    'Variable to deal with special cases of thematic vowel insertion
    Dim newThematic As String
    'Randomizes the visaul basic random number generator
    Randomize
    'Sets special cases of person and number for participles
    testNumber = 3 ' none
    testPerson = 4 ' none
    testTense = Int((3 - 1 + 1) * Rnd + 1) ' randomly generates the tense for participles (1-3) present, future, perfect
    
    If testConjugation = 4 Or testConjugation = 5 Then ' if the conjugation is 3rd-io or fourth then give special thematice
        newThematic = "ie"
    Else
        newThematic = VowelIndicative(testConjugation) ' else the normal indicative thematic
    End If
    'deals with deponent and defective verbs
    If VerbType = 2 Then ' if deponent
        testTense = Int((3 - 2 + 1) * Rnd + 2) 'limit tense to future or paerfect(i.e. those with passives)
        testVoice = 2 ' sets voice = passive
    ElseIf VerbType = 4 Then ' If defective limit the tense to present and future
        testTense = Int((2 - 1 + 1) * Rnd + 1)
    End If
    'Main descision engine
    Select Case testTense
        Case 1 'If present give pres endings and limit voice to active
339         If testVoice = 1 Then
                TestVerb = VerbPresStem(rndVerb) & newThematic & "ns,-ntis"
            Else
                TestVerb = VerbPresStem(rndVerb) & newThematic & "ns,-ntis"
                testVoice = 1
            End If
        Case 2 ' If future give future endings
            If testVoice = 1 Then
                TestVerb = VerbPartStem(rndVerb) & "urus -a -um"
            Else
                TestVerb = VerbPresStem(rndVerb) & newThematic & "ndus, -a -um"
            End If
        Case 3 ' if perfect limit voice to passive
360         If testVoice = 2 Then
                TestVerb = VerbPartStem(rndVerb) & "us -a -um"
            Else
                testVoice = 2
                TestVerb = VerbPartStem(rndVerb) & "us -a -um"
            End If
    End Select
    
    Select Case testTense
        Case 1 'present
            testTense = 1
        Case 3 ' perfect
            testTense = 4
        Case 2 ' future
            testTense = 3
    End Select
    
    
End Sub

Public Sub Infinitives()
    'gives and limits infinitive verb endings
    'Randomizes visual basics random number generator
    Randomize
    'Limits number and person to none
    testNumber = 3
    testPerson = 4
    'Randomly select a tense 1 - 3 for infinitives
    testTense = Int((3 - 1 + 1) * Rnd + 1)
    'Deals with deponenet and defective verbs
    If VerbType = 2 Then
        Do Until testTense <> 2
            testTense = Int((3 - 1 + 1) * Rnd + 1)
        Loop
            testVoice = 2
    ElseIf VerbType = 4 Then
            testTense = Int((2 - 1 + 1) * Rnd + 1)
    End If
    'main descision engine of infinitives
    Select Case testTense
        Case 1 ' If present
            If testVoice = 1 Then 'If active
                TestVerb = VerbInfinitive(rndVerb)
            Else 'If passive
                Select Case testConjugation
                    Case 1, 2, 5  ' if not thrid
                        TestVerb = VerbPresStem(rndVerb) & VowelIndicative(testConjugation) & "ri"
                    Case 3, 4 ' if Third
                        TestVerb = VerbPresStem(rndVerb) & "i"
                End Select
            End If
        Case 2 'If future
422         If testVoice = 1 Then ' if active
                TestVerb = VerbPartStem(rndVerb) & "urus(-a -um) esse"
            Else
                TestVerb = VerbPartStem(rndVerb) & "urus(-a -um) esse"
                testVoice = 1
            End If
        Case 3 'If Perfect
            If testVoice = 1 Then
                TestVerb = VerbPerfStem(rndVerb) + "isse"
            Else
                TestVerb = VerbPartStem(rndVerb) & "us(-a,-um) esse"
            End If
    End Select
    'Gives universal tense to special tense
    Select Case testTense
        Case 1 ' present
            testTense = 1
        Case 2 ' future
            testTense = 3
        Case 3 ' perfect
            testTense = 4
    End Select
    
End Sub
Public Sub TestVerbLevel()
    'Test the verb level for the selected verb and if verb does not match it generates another verb and if not match is found
    'Randomizes visaul basics random number generator
    Randomize
    'Declares useful variables (localized
    Dim atLeastOne As Boolean
    Dim good As Boolean
    Dim pos As Integer
    'Initializes variables
    atLeastOne = False
    good = False
    pos = 0
    ' Checks to see if there is any verb in the arrays which matches user's student level
    Do Until atLeastOne Or pos = verbCtr
        pos = pos + 1
        If VerbDifficulty(rndVerb) <= StudentLevel Then
            atLeastOne = True
        End If
    Loop
    'If there's is at least one verb
    If atLeastOne Then
        'Then check to see if the one that we currently have is either that one or another one
        Do Until good
            'If it is higher than the student's level then generate a new verb
            If VerbDifficulty(rndVerb) > StudentLevel Then
                rndVerb = Int((verbCtr - 1 + 1) * Rnd + 1)
            Else
            'If it does match, end the loop by setting good to true
                good = True
            End If
        Loop
    Else ' If there is not at least one display message, and end the current sub and pass on that the sub routine need to be stopped by setting the exitSub variable to true
        MsgBox "There are no verbs appropriate for your class level, please contact your administrator to remedy this situation"
        ExitSUB = True
        Exit Sub
    End If
    
End Sub

Public Sub TestFormLevel()
    'Generates all of the criteria for the forma and checks to see if it is appropriate for the level of the student
    'Randomizes the visual basics' random number generator
    Randomize
    'Declares useful varaibles
    Dim pos As Integer
    Dim tempType As Integer
    Dim Found As Boolean
    
477 'This is used for the case where the verb form does not match (in which case the program reinitializes and regenerates all criteria)
    'Generates randomly all the criteria for the verb
    rndVerb = Int((verbCtr - 1 + 1) * Rnd + 1)
    testVoice = Int((2 - 1 + 1) * Rnd + 1)
    testMood = Int((5 - 1 + 1) * Rnd + 1)
    testNumber = Int((2 - 1 + 1) * Rnd + 1)
    testPerson = Int((3 - 1 + 1) * Rnd + 1)
    'Checks the verb in order to see if it correct
    Call TestVerbLevel
    
    If ExitSUB Then
        Exit Sub
    End If
    'Gets the verb conjugation and the verb type based on the random verb information
    testConjugation = VerbConjugation(rndVerb)
    VerbType = VerbClass(rndVerb)
    'Checks what mood and enters the appropriate public subroutine to generate the ending and the testVerb
    If testMood = 1 Then
        tMoodName = "Indic."
        Call IndicativeSequence
    ElseIf testMood = 2 Then
        tMoodName = "Subj."
        Call SubjungtiveSequence
    ElseIf testMood = 3 Then
        tMoodName = "Imper."
        Call Imperatives
    ElseIf testMood = 4 Then
        tMoodName = "Part."
        Call Participles
    ElseIf testMood = 5 Then
        tMoodName = "Infin."
        Call Infinitives
    Else
        MsgBox "Your RndNumber Range is incorrect"
        ExitSUB = True
        Exit Sub
    End If
    'loops over the formClass array to check and see if the verb type is less than or greater than the student leve
    For pos = 1 To verbFormctr
        If VerbType = formClass(pos) Then
            If formClassLevel(pos) > StudentLevel Then ' If the formClassLevel of the specified verb type is greater than student level start selection process again
                GoTo 477
            End If
        End If
    Next pos
    
    'MsgBox "tense " & testTense & " mood " & testMood & " voice " & testVoice ' debugging message
    'Initializes varaibles
    pos = 0
    Found = False
    'loops over 3 arrays simulataneously to check if the verb form is appropriate for the user (match and stop to find the position of the form
    Do Until Found Or pos = verbFormctr
        pos = pos + 1
        If testTense = formTense(pos) And testMood = formMood(pos) And testVoice = formVoice(pos) Then
            Found = True
        End If
    Loop
    
    'Checks the formClassleve agains the studentlevel and if they are incompatible restarts the entire process
    If Found Then
        If formClassLevel(pos) > StudentLevel Then GoTo 477
    Else
        MsgBox "the testForm Algorithm is not working, try something different" 'diagnostic message
    End If
 
    
End Sub

Private Sub cmdEnd_Click()
    'Ends the testing session
    'Enables and disables and makes visible and invisible appropriate buttons
    cmdEnd.Visible = False
    cmdStart.Visible = True
    cmdLogOut.Enabled = True
    cmdReturn.Enabled = True
    cmdSubmit.Enabled = False
    'Clears the verb and principle parts picBoxes
    picVerb.Cls
    picPrinciple.Cls
    'Calculates the grade of the user and stores it in global arrays for use after logout
    Call CalculateGrade(NumCorrect, NumWrong, maxGrade)
           
End Sub

Private Sub cmdLogOut_Click()
    'Logs out the user
    frmVerbTest.Hide
    Call LogOut
End Sub

Private Sub cmdReturn_Click()
    ' returns to the student options pane
    frmVerbTest.Hide
    frmOptionsPage.Show
End Sub

Private Sub cmdStart_Click()
    'Starts the testing session
    'Randomizes the visual basic random number generator
    Randomize
    'Enables disables, makes visible and invivisible appropriate buttons
    cmdSubmit.Enabled = True
    cmdStart.Visible = False
    cmdLogOut.Enabled = False
    cmdEnd.Visible = True
    cmdReturn.Enabled = False
    'Initializes grading varaibles
    maxGrade = 0
    NumWrong = 0
    NumCorrect = 0
    ExitSUB = False
    'Initializes picBoxes
    picGrade.Cls
    picGrade.Print NumCorrect & "/"; maxGrade
    picVerbsTested.Cls
    picVerbsTested.Print "Verb"; Tab(30); "Tested Form"; Tab(68); "Answered Form"; Tab(103); "Correct"
    picVerbsTested.Print Tab(30); "(Conjugation tense, voice, mood; number, person)"
    picVerbsTested.Print "**************************************************************************************************************************************************"
    
    'calls the main testing function of the program
    Call TestFormLevel
    'Checks to see if it needs to exit the subroutine, and does so if need be
    If ExitSUB Then
        Exit Sub
    End If
    
    If VerbType = 2 Then
        testVoice = 2
    End If
    
    'Prints the generated verb and appropriate principle parts
    picVerb.Print TestVerb
    picPrinciple.Print VerbPrincipleParts(rndVerb)
    
    'Gives names to the number with which we were working
    Select Case testTense
        Case 1
            tTenseName = "Pres."
        Case 2
            tTenseName = "Imperf."
        Case 3
            tTenseName = "Fut."
        Case 4
            tTenseName = "Perf."
        Case 5
            tTenseName = "PluPerf."
        Case 6
            tTenseName = "FutPerf."
    End Select
    
    Select Case testVoice
        Case 1
            tVoiceName = "Act."
        Case 2
            tVoiceName = "Pass."
    End Select
    
    Select Case testNumber
        Case 1
            tNumName = "S."
        Case 2
            tNumName = "P."
        Case 3
            tNumName = ""
    End Select
        
    Select Case testConjugation
        Case 1
            tConjName = "1st"
        Case 2
            tConjName = "2nd"
        Case 3
            tConjName = "3rd"
        Case 4
            tConjName = "3rd-IO"
        Case 5
            tConjName = "4th"
    End Select
    
    Select Case testPerson
        Case 1
            tPersonName = "1st,"
        Case 2
            tPersonName = "2nd,"
        Case 3
            tPersonName = "3rd,"
        Case 4
            tPersonName = ""
    End Select
    
    previousVerb = TestVerb
            
End Sub

Private Sub cmdSubmit_Click()
    'Submits the answer for verification and generates new verb form
    
    Dim matched As Boolean
    matched = True
    'Randomizes the visual basic random number generator
    Randomize
    'Initializes the exit sub boolean
    ExitSUB = False
    'Gets the answer data from the user
    Call GetAnswerData
    'Exits sub if user missed any entry
    If ExitSUB Then
        Exit Sub
    End If
    'Test the answer and prints appropriate entry on picVerbsTested (cf. TestAnswer below for details)
    Call TestAnswer
    'Increments max grade
    maxGrade = maxGrade + 1
    
    'Gives the user infor about correct or incorrect
    picGrade.Cls
    picGrade.Print NumCorrect & "/" & maxGrade
    'Generates new verb
    Call TestFormLevel
    'Makes sure that the last verbForm is not the same as the previous
    Do Until Not matched
        If TestVerb = previousVerb Then
            Call TestFormLevel
        Else
            previousVerb = TestVerb
            matched = False
        End If
    Loop
    'Checks to see if it needs to exit the sub
    If VerbType = 2 Then
        testVoice = 2
    End If
    
    If ExitSUB Then
        Exit Sub
    End If
    'reinitializes picBoxes and prints appropriate output
    picVerb.Cls
    picVerb.Print TestVerb
    picPrinciple.Cls
    picPrinciple.Print VerbPrincipleParts(rndVerb)
    'Gives names to the nubmer with which we are working for user output
    Select Case testTense
        Case 1
            tTenseName = "Pres."
        Case 2
            tTenseName = "Imperf."
        Case 3
            tTenseName = "Fut."
        Case 4
            tTenseName = "Perf."
        Case 5
            tTenseName = "PluPerf."
        Case 6
            tTenseName = "FutPerf."
    End Select
    
    Select Case testVoice
        Case 1
            tVoiceName = "Act."
        Case 2
            tVoiceName = "Pass."
    End Select
    
    Select Case testNumber
        Case 1
            tNumName = "S."
        Case 2
            tNumName = "P."
        Case 3
            tNumName = ""
    End Select
        
    Select Case testConjugation
        Case 1
            tConjName = "1st"
        Case 2
            tConjName = "2nd"
        Case 3
            tConjName = "3rd"
        Case 4
            tConjName = "3rd-IO"
        Case 5
            tConjName = "4th"
    End Select
    
    Select Case testPerson
        Case 1
            tPersonName = "1st,"
        Case 2
            tPersonName = "2nd,"
        Case 3
            tPersonName = "3rd,"
        Case 4
            tPersonName = ""
    End Select
    
End Sub

Public Sub TestAnswer()
'Checks the answer against the test criteria and either increments numCorrect or Numwrong and prints appropriate message
    If maxGrade = 35 Or maxGrade = 70 Or maxGrade = 105 Or maxGrade = 140 Or maxGrade = 175 Then
        picVerbsTested.Cls
        picVerbsTested.Print "Verb"; Tab(30); "Tested Form"; Tab(68); "Answered Form"; Tab(103); "Correct"
        picVerbsTested.Print Tab(30); "(Conjugation tense, voice, mood; number, person)"
        picVerbsTested.Print "**************************************************************************************************************************************************"
    ElseIf maxGrade = 210 Then
        'Ends the testing session
        'Enables and disables and makes visible and invisible appropriate buttons
        cmdEnd.Visible = False
        cmdStart.Visible = True
        cmdLogOut.Enabled = True
        cmdReturn.Enabled = True
        cmdSubmit.Enabled = False
        'Clears the verb and principle parts picBoxes
        picVerb.Cls
        picPrinciple.Cls
        'Calculates the grade of the user and stores it in global arrays for use after logout
        Call CalculateGrade(NumCorrect, NumWrong, maxGrade)
        MsgBox "You've been at this for a long time, take a break and then come back to us. A well-rested brain is a happy brain!"
    End If
        
    If testTense = answerTense And testMood = answerMood And testVoice = answerVoice And testNumber = answerNumber And testPerson = answerPerson And testConjugation = answerConjugation Then
        picVerbsTested.Print TestVerb; Tab(30); tConjName & " " & tTenseName & ", " & tVoiceName & ", " & tMoodName & "; " & tPersonName & " " & tNumName; Tab(68); AConjName & " " & ATenseName & ", " & AVoiceName & ", " & AMoodName & "; " & APersonName & " " & ANumName; Tab(103); "YES"
        NumCorrect = NumCorrect + 1
    Else
        picVerbsTested.Print TestVerb; Tab(30); tConjName & " " & tTenseName & ", " & tVoiceName & ", " & tMoodName & "; " & tPersonName & " " & tNumName; Tab(68); AConjName & " " & ATenseName & ", " & AVoiceName & ", " & AMoodName & "; " & APersonName & " " & ANumName; Tab(103); "NO"
        NumWrong = NumWrong + 1
    End If
    
    
End Sub
Private Sub Form_Load()
    'Does various actions upon form Load
    'Randomizes the visual basic random number generator
    Randomize
    'declares a pos varaible for reading the text files into arrays
    Dim pos As Integer
    'reads the thematix vowels
    Open App.Path & "\data\ThematicVowels.txt" For Input As #1
        
        For pos = 1 To 5
            Input #1, VowelConjugation(pos), VowelIndicative(pos), VowelSubjungtive(pos)
        Next pos
    Close #1
    'Reads the indicative sequence endings
    Open App.Path & "\Data\IndicativeSequence.txt" For Input As #2
        IendingCtr = 0
        Do Until EOF(2)
            IendingCtr = IendingCtr + 1
            Input #2, IVerbFormName(IendingCtr), IFirstS(IendingCtr), ISecondS(IendingCtr), IThirdS(IendingCtr), IFirstP(IendingCtr), ISecondP(IendingCtr), IThirdP(IendingCtr)
        Loop
    Close #2
    'reads the subjungtive sequence endings
    Open App.Path & "\Data\SubjungtiveSequence.txt" For Input As #3
    
        For pos = 1 To 8
            Input #3, SVerbFormName(pos), SFirstS(pos), SSecondS(pos), SThirdS(pos), SFirstP(pos), SSecondP(pos), SThirdP(pos)
        Next pos
        
    Close #3
    'reads the verb form level data from test files
    Open App.Path & "\Data\VerbFormsByClass.txt" For Input As #4
        verbFormctr = 0
        Do Until EOF(4)
            verbFormctr = verbFormctr + 1
            Input #4, verbFormLevel(verbFormctr), formClassLevel(verbFormctr), formTense(verbFormctr), formMood(verbFormctr), formVoice(verbFormctr), formClass(verbFormctr)
        Loop
    Close #4
    
End Sub

