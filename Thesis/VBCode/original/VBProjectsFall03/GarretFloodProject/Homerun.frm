VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080FF80&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "CLEAR"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Click here to find out who the players are that might break his homerun record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      TabIndex        =   6
      Top             =   4200
      Width           =   3135
   End
   Begin VB.PictureBox picbox 
      Height          =   975
      Left            =   2040
      ScaleHeight     =   915
      ScaleWidth      =   10515
      TabIndex        =   5
      Top             =   2760
      Width           =   10575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click here to see Hank Aaron's career stats"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.PictureBox Hankbox 
      Height          =   1215
      Left            =   6960
      ScaleHeight     =   1155
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Note:  Hank Aaron's age was his last year of playing baseball."
      Height          =   615
      Left            =   5640
      TabIndex        =   8
      Top             =   6120
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Career Homeruns:  755"
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Hank Aaron"
      Height          =   255
      Left            =   6960
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "WHO WILL BREAK HANK AARON'S       CAREER HOMERUN RECORD?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:  Project1 (Homerun.vbp)
'Form Name:  Form1 (Homerun.frm)
'Date Written:  Oct 28th, 2003
'Overall purpose of program:  To let the user find out information about
'the best homerun hitters at the present time and to find out what
'their chances are of breaking the all time homerun record held
'by Hank Aaron.
'Purpose of form:  To let the user see Hank Aaron's batting statistics.
'Then to go on to find out information on other players statistics.

'Option Explict makes the programmer declare all variables on the form.
Option Explicit
Dim YearsPlayed(1 To 12), AtBats(1 To 12), Hits(1 To 12), Homeruns(1 To 12), BattingAvg(1 To 12), Age(1 To 12) As Single
    Dim Player(1 To 12) As String
    Dim X As Integer

Private Sub Command1_Click()

'This command lets you see Hank Aaron's statistics
    X = 0
    X = X + 1
        picbox.Print "Player"; Tab(20); "Years Played"; Tab(40); "At Bats"; Tab(60); "Hits"; Tab(80); "HOMERUNS"; Tab(100); "Batting Avg."; Tab(120); "Age"
        picbox.Print "****************************************************************************************************************************************************************"
    Open PATH & "stats.txt" For Input As #1
        Input #1, Player(X), YearsPlayed(X), AtBats(X), Hits(X), Homeruns(X), BattingAvg(X), Age(X)
         picbox.Print Player(1); Tab(20); YearsPlayed(1); Tab(40); AtBats(1); Tab(60); Hits(1); Tab(80); Homeruns(1); Tab(100); BattingAvg(1); Tab(120); Age(1)
    'Hank Aaron is the first player listed on the array.
    Close #1
    
End Sub

Private Sub Command2_Click()
'This lets you go from form one to form two
    Form2.Show
    Form1.Hide
End Sub

Private Sub Command3_Click()
'clears picture box
    picbox.Cls
End Sub

Private Sub Form_Load()

'Lets you see a picture of Hank Aaron
    Hankbox.Picture = LoadPicture(PATH & "Hank.jpg")
End Sub
