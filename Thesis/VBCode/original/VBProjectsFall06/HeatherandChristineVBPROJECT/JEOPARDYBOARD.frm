VERSION 5.00
Begin VB.Form frmJeopardyBoard 
   BackColor       =   &H00400000&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Kristen ITC"
      Size            =   27.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTimer 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   885
      Left            =   3240
      TabIndex        =   37
      Top             =   9600
      Width           =   3135
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   885
      Left            =   12120
      TabIndex        =   36
      Text            =   " "
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2760
      Top             =   9000
   End
   Begin VB.Frame frmScoring 
      BackColor       =   &H00FFFF00&
      Caption         =   "Total Score"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   3240
      TabIndex        =   35
      Top             =   8880
      Width           =   3135
   End
   Begin VB.CommandButton cmdH1000 
      BackColor       =   &H80000009&
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdH800 
      BackColor       =   &H80000009&
      Caption         =   "800"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdH400 
      BackColor       =   &H80000009&
      Caption         =   "400"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdH200 
      BackColor       =   &H80000009&
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdEndGame 
      BackColor       =   &H00FFFF80&
      Caption         =   "End Game"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8760
      Width           =   2175
   End
   Begin VB.PictureBox picScore 
      BackColor       =   &H00FF0000&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   6360
      ScaleHeight     =   1395
      ScaleWidth      =   3075
      TabIndex        =   29
      Top             =   8520
      Width           =   3135
   End
   Begin VB.PictureBox picGender 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   12120
      ScaleHeight     =   3555
      ScaleWidth      =   2835
      TabIndex        =   28
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton cmdStartNew 
      BackColor       =   &H00FFFF80&
      Caption         =   "Start A New Game"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12480
      MaskColor       =   &H80000012&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7440
      Width           =   2175
   End
   Begin VB.Frame frmTitle 
      BackColor       =   &H00000000&
      Caption         =   "LETS PLAY JEOPARDY!"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   1575
      Left            =   720
      TabIndex        =   26
      Top             =   240
      Width           =   13815
   End
   Begin VB.CommandButton cmdArts1000 
      BackColor       =   &H8000000E&
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdArts800 
      BackColor       =   &H8000000E&
      Caption         =   "800"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdArts600 
      BackColor       =   &H8000000E&
      Caption         =   "600"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdArts400 
      BackColor       =   &H8000000E&
      Caption         =   "400"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdArts200 
      BackColor       =   &H8000000E&
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSL1000 
      BackColor       =   &H8000000E&
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSL800 
      BackColor       =   &H8000000E&
      Caption         =   "800"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdSL600 
      BackColor       =   &H8000000E&
      Caption         =   "600"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSL400 
      BackColor       =   &H8000000E&
      Caption         =   "400"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdSL200 
      BackColor       =   &H8000000E&
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdA1000 
      BackColor       =   &H8000000E&
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdA800 
      BackColor       =   &H8000000E&
      Caption         =   "800"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdA600 
      BackColor       =   &H8000000E&
      Caption         =   "600"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdA400 
      BackColor       =   &H8000000E&
      Caption         =   "400"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdA200 
      BackColor       =   &H8000000E&
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdS1000 
      BackColor       =   &H8000000E&
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdS800 
      BackColor       =   &H8000000E&
      Caption         =   "800"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdS600 
      BackColor       =   &H8000000E&
      Caption         =   "600"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdS400 
      BackColor       =   &H8000000E&
      Caption         =   "400"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdS200 
      BackColor       =   &H8000000E&
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdH600 
      BackColor       =   &H8000000E&
      Caption         =   "600"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame frmArts 
      BackColor       =   &H00000000&
      Caption         =   "ARTS"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   10200
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Frame frmStudentLife 
      BackColor       =   &H00000000&
      Caption         =   "STUDENT LIFE"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   7320
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Frame frmAcademics 
      BackColor       =   &H00000000&
      Caption         =   "ACADEMICS"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Frame frmSports 
      BackColor       =   &H00000000&
      Caption         =   "SPORTS"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Frame frmHistory 
      BackColor       =   &H00000000&
      Caption         =   "HISTORY"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   11400
      Left            =   0
      Picture         =   "JEOPARDY BOARD.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15240
   End
End
Attribute VB_Name = "frmJeopardyBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jeopardy Board: On this form the user will begin playing Jeopardy.
'                The user will select a value and a question will pop up.
Option Explicit
Dim InputAnswer As String
Dim Sum As Integer
Dim QTime As Integer
Dim Limit As Integer




Private Sub cmdA1000_Click()
    InputAnswer = InputBox("If you need help with your math homework, where can you go?")
    If InputAnswer = "Math Skills Center" Or InputAnswer = "math skills center" Or InputAnswer = "Math skills center" Then
        MsgBox "That is the correct answer!"
        cmdA1000.Visible = False
        Sum = Sum + 1000
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        cmdA1000.Visible = False
    End If
    picScore.Cls
    picScore.Print Sum
End Sub

Private Sub cmdA200_Click()
    InputAnswer = InputBox("How many majors are offered at CSB/SJU?")
    If InputAnswer = "40" Then
        MsgBox "That is the correct answer!"
        cmdA200.Visible = False
        Sum = Sum + 200
     Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        cmdA200.Visible = False
    End If
    picScore.Cls
    picScore.Print Sum
End Sub

Private Sub cmdA400_Click()
    InputAnswer = InputBox("What percent of students graduate in 4 years?")
    If InputAnswer = "90%" Or InputAnswer = "90" Then
        MsgBox "That is the correct answer!"
        cmdA400.Visible = False
        Sum = Sum + 400
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        cmdA400.Visible = False
    End If
    picScore.Cls
    picScore.Print Sum
End Sub

Private Sub cmdA600_Click()
    MsgBox "DAILY DOUBLE!!!!"
    InputAnswer = InputBox("What is the professor/student ratio?")
    If InputAnswer = "13:1" Or InputAnswer = "13 to 1" Then
        MsgBox "That is the correct answer!"
        cmdA600.Visible = False
        Sum = Sum + 1200
        
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        cmdA600.Visible = False
        Sum = Sum + 0
    End If
    picScore.Cls
    picScore.Print Sum
End Sub

Private Sub cmdA800_Click()
    InputAnswer = InputBox("How many semester-long study abroad programs are offered?")
    If InputAnswer = "17" Then
        MsgBox "That is the correct answer!"
        cmdA800.Visible = False
        Sum = Sum + 800
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        cmdA800.Visible = False
    End If
    picScore.Cls
    picScore.Print Sum
End Sub

Private Sub cmdArts1000_Click()
    InputAnswer = InputBox("In what year did the BAC open?")
    If InputAnswer = "1963" Then
        MsgBox "That is the correct answer!"
        cmdArts1000.Visible = False
        Sum = Sum + 1000
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        cmdArts1000.Visible = False
    End If
    picScore.Cls
    picScore.Print Sum
End Sub

Private Sub cmdArts200_Click()
    InputAnswer = InputBox("What does the BAC stand for")
    If InputAnswer = "benedicta arts center" Or InputAnswer = "Benedicta arts center" Or InputAnswer = "Benedicta Arts Center" Then
        MsgBox "That is the correct answer!"
        cmdArts200.Visible = False
        Sum = Sum + 200
        picScore.Cls
        picScore.Print Sum
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        Sum = Sum + 0
        picScore.Cls
        picScore.Print Sum
        cmdArts200.Visible = False
    End If
End Sub

Private Sub cmdArts400_Click()
    InputAnswer = InputBox("How many FAE events are CSB/SJU student to attend?")
    If InputAnswer = "7" Then
        MsgBox "That is the correct answer!"
        cmdArts400.Visible = False
        Sum = Sum + 400
        picScore.Cls
        picScore.Print Sum
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        Sum = Sum + 0
        picScore.Cls
        picScore.Print Sum
        cmdArts400.Visible = False
    End If
End Sub

Private Sub cmdArts600_Click()
    InputAnswer = InputBox("How many art majors are offered?")
    If InputAnswer = "2" Then
        MsgBox "That is the correct answer!"
        cmdArts600.Visible = False
        Sum = Sum + 600
        picScore.Cls
        picScore.Print Sum
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        Sum = Sum + 0
        picScore.Cls
        picScore.Print Sum
        cmdArts600.Visible = False
    End If
End Sub

Private Sub cmdArts800_Click()
    MsgBox "DAILY DOUBLE!!!!!"
    InputAnswer = InputBox("What are the 3 components of the Fine Arts Division? (Please use commas to seperate your answer)")
    If InputAnswer = "Art, Music, Theater" Or InputAnswer = "art, music, theater" Or InputAnswer = "music, art, theater" Or InputAnswer = "music, theater, art" Or InputAnswer = "art, music, theater" Or InputAnswer = "art, theater, music" Or InputAnswer = "theater, music, art" Or InputAnswer = "theater, art, music" Then
        MsgBox "That is the correct answer!"
        cmdArts800.Visible = False
        Sum = Sum + 1600
        picScore.Cls
        picScore.Print Sum
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        Sum = Sum + 0
        picScore.Cls
        picScore.Print Sum
        cmdArts800.Visible = False
    End If
End Sub

Private Sub cmdEndGame_Click()
'The End Game button sends the player to the last form which is the check the
'   player would receive with the amount of money they earned while playing
'   the game. Once the player presses this button they will not be able to begin
'   a new game.

    frmJeopardyBoard.Hide
    frmMoney.Show
    frmMoney.txtPlayerMoney.Text = PlayerName
    frmMoney.txtGrandTotal.Text = FormatCurrency(Sum)
    txtTimer.Text = Time
    frmMoney.txtEndingTime.Text = "Your Ending Time is, " & Time
End Sub

Private Sub cmdH1000_Click()
    InputAnswer = InputBox("Who is SJU's current President?")
    If InputAnswer = "Dietrich Reinhart" Or InputAnswer = "Brother Dietrich Reinhart" Or InputAnswer = "Reinhart" Then
        MsgBox "That is the correct answer!"
        cmdH1000.Visible = False
        Sum = Sum + 1000
        picScore.Cls
        picScore.Print Sum
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        Sum = Sum + 0
        picScore.Cls
        picScore.Print Sum
        cmdH1000.Visible = False
    End If
End Sub

Private Sub cmdH200_Click()
' Each of the numbered buttons allows the player to view the question for that
'   amount of points. The points will automatically be added to the sum if the
'   player answers the question correctly!

    InputAnswer = InputBox("What was the first building at CSB?")
          
    If InputAnswer = "Main" Or InputAnswer = "main" Then
        MsgBox "That is the correct answer!"
        cmdH200.Visible = False
        Sum = Sum + 200
        picScore.Cls
        picScore.Print Sum
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        Sum = Sum + 0
        picScore.Cls
        picScore.Print Sum
        cmdH200.Visible = False
    End If
       
End Sub



Private Sub cmdH400_Click()
    InputAnswer = InputBox("What year did CSBSJU combined classes?")
    If InputAnswer = "1963" Then
        MsgBox "That is the correct answer!"
        cmdH400.Visible = False
        Sum = Sum + 400
        picScore.Cls
        picScore.Print Sum
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        Sum = Sum + 0
        picScore.Cls
        picScore.Print Sum
        cmdH400.Visible = False
    End If
        

End Sub





Private Sub cmdH600_Click()
    InputAnswer = InputBox("What year was CSB found?")
    If InputAnswer = "1913" Then
        MsgBox "That is the correct answer!"
        cmdH600.Visible = False
        Sum = Sum + 600
        picScore.Cls
        picScore.Print Sum
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        Sum = Sum + 0
        picScore.Cls
        picScore.Print Sum
        cmdH600.Visible = False
    End If
End Sub



Private Sub cmdH800_Click()
    InputAnswer = InputBox("Rounded to the nearest 10,000, How many alumni does CSBSJU have?")
    If InputAnswer = "40,000" Or InputAnswer = "41,000" Then
        MsgBox "That is the correct answer!"
        cmdH800.Visible = False
        Sum = Sum + 800
        picScore.Cls
        picScore.Print Sum
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        Sum = Sum + 0
        picScore.Cls
        picScore.Print Sum
        cmdH800.Visible = False
    End If
End Sub

Private Sub cmdH900_Click()
    cmdH
End Sub

Private Sub cmdS1000_Click()
    InputAnswer = InputBox("Who is CSB's athletic director?")
    If InputAnswer = "Carol Howe-Veenstra" Or InputAnswer = "Carol" Then
        MsgBox "That is the correct answer!"
        cmdS1000.Visible = False
        Sum = Sum + 1000
        picScore.Cls
        picScore.Print Sum
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        Sum = Sum + 0
        picScore.Cls
        picScore.Print Sum
        cmdS1000.Visible = False
    End If
    
End Sub

Private Sub cmdS200_Click()
    InputAnswer = InputBox("How many sports are offered at CSB?")
    If InputAnswer = "11" Then
        MsgBox "That is the correct answer!"
        cmdS200.Visible = False
        Sum = Sum + 200
        picScore.Cls
        picScore.Print Sum
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        Sum = Sum + 0
        picScore.Cls
        picScore.Print Sum
        cmdS200.Visible = False
    End If
End Sub

Private Sub cmdS400_Click()
    InputAnswer = InputBox("How many sports are offered at SJU?")
    If InputAnswer = "12" Then
        MsgBox "That is the correct answer!"
        cmdS400.Visible = False
        Sum = Sum + 400
        picScore.Cls
        picScore.Print Sum
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        Sum = Sum + 0
        picScore.Cls
        picScore.Print Sum
        cmdS400.Visible = False
    End If
End Sub

Private Sub cmdS600_Click()
    InputAnswer = InputBox("Who is the SJU football coach?")
    If InputAnswer = "John Gagliardi" Or InputAnswer = "John G" Or InputAnswer = "John" Then
        MsgBox "That is the correct answer!"
        cmdS600.Visible = False
        Sum = Sum + 600
        picScore.Cls
        picScore.Print Sum
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        Sum = Sum + 0
        picScore.Cls
        picScore.Print Sum
        cmdS600.Visible = False
    End If
    
End Sub

Private Sub cmdS800_Click()
    InputAnswer = InputBox("What is SJU's mascot?")
    If InputAnswer = "rat" Or InputAnswer = "Rat" Then
        MsgBox "That is the correct answer!"
        cmdS800.Visible = False
        Sum = Sum + 800
        picScore.Cls
        picScore.Print Sum
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        Sum = Sum + 0
        picScore.Cls
        picScore.Print Sum
        cmdS800.Visible = False
    End If
    
End Sub

Private Sub cmdSL1000_Click()
    InputAnswer = InputBox("What does the program R.A.D stand for?")
    If InputAnswer = "Rape aggression defense system" Or InputAnswer = "rape aggression defense system" Then
        MsgBox "That is the correct answer!"
        cmdSL1000.Visible = False
        Sum = Sum + 1000
        picScore.Cls
        picScore.Print Sum
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        Sum = Sum + 0
        picScore.Cls
        picScore.Print Sum
        cmdSL1000.Visible = False
    End If
End Sub

Private Sub cmdSL200_Click()
    InputAnswer = InputBox("True or False, are students required to live on campus their first year only?")
    If InputAnswer = "false" Or InputAnswer = "False" Then
        MsgBox "That is the correct answer!"
        cmdSL200.Visible = False
        Sum = Sum + 200
        picScore.Cls
        picScore.Print Sum
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        Sum = Sum + 0
        picScore.Cls
        picScore.Print Sum
        cmdSL200.Visible = False
    End If
End Sub

Private Sub cmdSL400_Click()
    InputAnswer = InputBox("How many clubs and organizations are there offered?")
    If InputAnswer = "65" Then
        MsgBox "That is the correct answer!"
        cmdSL400.Visible = False
        Sum = Sum + 400
        picScore.Cls
        picScore.Print Sum
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        Sum = Sum + 0
        picScore.Cls
        picScore.Print Sum
        cmdSL400.Visible = False
    End If
End Sub

Private Sub cmdSL600_Click()
    InputAnswer = InputBox("What does SIFE stand for")
    If InputAnswer = "Students in free enterprise" Or InputAnswer = "students in free enterprise" Or InputAnswer = "Students in Free Enterprise" Then
        MsgBox "That is the correct answer!"
        cmdSL600.Visible = False
        Sum = Sum + 600
        picScore.Cls
        picScore.Print Sum
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        Sum = Sum + 0
        picScore.Cls
        picScore.Print Sum
        cmdSL600.Visible = False
    End If
End Sub

Private Sub cmdSL800_Click()
    InputAnswer = InputBox("What percent of CSB/SJU is women?")
    If InputAnswer = "52%" Or InputAnswer = "52" Then
        MsgBox "That is the correct answer!"
        cmdSL800.Visible = False
        Sum = Sum + 800
        picScore.Cls
        picScore.Print Sum
    Else
        MsgBox "Sorry!  That is not the answer we are looking for!"
        Sum = Sum + 0
        picScore.Cls
        picScore.Print Sum
        cmdSL800.Visible = False
    End If
End Sub

Private Sub cmdStartNew_Click()
' This button simply allows the user to return to the first form and begin a
'   new game if they would like.

    frmJeopardyBoard.Hide
    frmGamePage.Show
End Sub


Private Sub Form_Load()
    frmJeopardyBoard.Caption = "Welcome PlayerName"
    Limit = 5
    Timer1.Enabled = True
End Sub


Private Sub Timer1_Timer()
' The timer will show the player the time as they are playing the game, if he or she
'   needs to be done at a certain time or they want to time themselves for how long
'   it will take them, this feature will be very convenient for he or she.
    QTime = QTime + 1
    txtTimer.Text = Time
    
End Sub


