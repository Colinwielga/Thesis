VERSION 5.00
Begin VB.Form frmHighScores 
   BackColor       =   &H00FF0000&
   Caption         =   "High Scores"
   ClientHeight    =   5160
   ClientLeft      =   2670
   ClientTop       =   3075
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   10425
   Begin VB.CommandButton cmdLoad2 
      BackColor       =   &H000000FF&
      Caption         =   "Click here to load most recent data!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Click here!"
      Top             =   120
      Width           =   10095
   End
   Begin VB.PictureBox Picture2 
      Height          =   1815
      Left            =   5040
      Picture         =   "frmHighScores.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   1395
      TabIndex        =   21
      Top             =   2280
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   4920
      Picture         =   "frmHighScores.frx":0D26
      ScaleHeight     =   1275
      ScaleWidth      =   1635
      TabIndex        =   20
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmdFindPlayer 
      BackColor       =   &H000000FF&
      Caption         =   "Find Player"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtEnterPlayer 
      Height          =   285
      Left            =   3360
      TabIndex        =   15
      Text            =   "Enter Player"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdFindHigher 
      BackColor       =   &H000000FF&
      Caption         =   "Find Higher"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtEnterScore 
      Height          =   285
      Left            =   3360
      TabIndex        =   4
      Text            =   "Enter Score"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Return to Welcome Screen"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   7935
   End
   Begin VB.CommandButton cmdSortNum 
      BackColor       =   &H000000FF&
      Caption         =   "Sort by Score (highest to lowest)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Sort Numerically"
      Top             =   810
      Width           =   3615
   End
   Begin VB.CommandButton cmdSortAlpha 
      BackColor       =   &H000000FF&
      Caption         =   "Sort by Player (a - z)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Sort Alphabetically"
      Top             =   1200
      Width           =   3615
   End
   Begin VB.PictureBox picHighScores 
      BackColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   6840
      ScaleHeight     =   4035
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   0
      TabIndex        =   18
      Top             =   4440
      Width           =   10215
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   0
      TabIndex        =   17
      Top             =   5040
      Width           =   10335
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF0000&
      Caption         =   "Input a player and click Find Player to display the player's score"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   14
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF0000&
      Caption         =   "Input a score and click Find Higher to display ALL players with scores higher than that"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   10200
      TabIndex        =   11
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   4680
      TabIndex        =   10
      Top             =   -120
      Width           =   135
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   0
      TabIndex        =   9
      Top             =   1680
      Width           =   4695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Searching Options"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Sorting Options"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "frmHighScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'declare form level variables
Dim Names(1 To 10) As String, HighScores(1 To 10) As Integer
Dim Pass As Integer, Pos As Integer, CTRHighScores As Integer

Private Sub cmdBack_Click()
    'hide high scores form
    frmHighScores.Hide
    
    'show welcome form
    frmWelcome.Show
End Sub

Private Sub cmdFindHigher_Click()
    'declare subroutine variables
    Dim InputScore As Single, matches As Boolean
    
    'ensure that search criteria is an integer since all scores are integers
    InputScore = Int(txtEnterScore.Text)
    
    'clear screen of previous results and reformat for new results
    picHighScores.Cls
    picHighScores.Print "Players with a score greater than " & FormatCurrency(InputScore, 0) & ":"
    picHighScores.Print
    picHighScores.Print "Player"; Tab(20); "Score"
    picHighScores.Print
    picHighScores.Print "**********************************************************************"
    
    'conduct exhaustive search, printing as you go
    For Pos = 1 To CTRHighScores
        If InputScore < HighScores(Pos) Then
            picHighScores.Print Names(Pos); Tab(20); FormatCurrency(HighScores(Pos), 0)
            matches = True
        End If
    Next Pos
    
    'display appropriate results for no matches
    If matches = False Then
        picHighScores.Print "Sorry!  No players scored higher than that!"
    End If
End Sub

Private Sub cmdFindPlayer_Click()
    'declare subroutine variables
    Dim InputPlayer As String, Found As Boolean
    
    'assign user input a variable name
    InputPlayer = txtEnterPlayer.Text
    
    'initialize variables
    Found = False
    Pos = 0
    
    'conduct match and stop search
    Do While (Found = False And Pos < CTRHighScores)
        Pos = Pos + 1
        If LCase(Names(Pos)) = LCase(InputPlayer) Then
            Found = True
        End If
    Loop
    
    'clear previous results
    picHighScores.Cls
       
    'print relevant results
    If Found = True Then
        picHighScores.Print Names(Pos) & " has a score of " & FormatCurrency(HighScores(Pos), 0) & "."
    Else
        picHighScores.Print "Sorry!  No players with that name!"
    End If
End Sub

Private Sub cmdLoad2_Click()
    'open high score data file and load into an array
    Open App.Path & "\HighScores.txt" For Input As #2
    CTRHighScores = 0
    Do Until EOF(2)
        CTRHighScores = CTRHighScores + 1
        Input #2, Names(CTRHighScores), HighScores(CTRHighScores)
    Loop
    Close #2
    
    'manually enter most recent user into high score list
    CTRHighScores = CTRHighScores + 1
    Names(CTRHighScores) = Player
    HighScores(CTRHighScores) = Score
    
    'hide command buttom
    cmdLoad2.Visible = False
End Sub

Private Sub cmdSortAlpha_Click()
    'declare  subroutine variables
    Dim Temp As String, Temp2 As Integer
    
    'clear screen of previous results
    picHighScores.Cls
    
    ' use the bubble sort
    For Pass = 1 To (CTRHighScores - 1)
        For Pos = 1 To CTRHighScores - Pass
            If Names(Pos) > Names(Pos + 1) Then
                Temp = Names(Pos)
                Names(Pos) = Names(Pos + 1)
                Names(Pos + 1) = Temp
                Temp2 = HighScores(Pos)
                HighScores(Pos) = HighScores(Pos + 1)
                HighScores(Pos + 1) = Temp2
            End If
        Next Pos
    Next Pass
    
    'print the sorted list
    picHighScores.Print "Recent High Scores (a - z):"
    picHighScores.Print
    picHighScores.Print "Rank"; Tab(10); "Player"; Tab(30); "Score"
    picHighScores.Print
    picHighScores.Print "**********************************************************************"
    
    For Increment = 1 To 10
        picHighScores.Print FormatNumber(Increment, 0) & "."; Tab(10); Names(Increment); Tab(30); FormatCurrency(HighScores(Increment), 0)
    Next Increment
    
End Sub

Private Sub cmdSortNum_Click()
    'declare  subroutine variables
    Dim Temp3 As Integer, Temp4 As String
    
    'clear screen of previous results
    picHighScores.Cls
    
    ' use the bubble sort
    For Pass = 1 To (CTRHighScores - 1)
        For Pos = 1 To (CTRHighScores - Pass)
            If HighScores(Pos) < HighScores(Pos + 1) Then
                Temp3 = HighScores(Pos)
                HighScores(Pos) = HighScores(Pos + 1)
                HighScores(Pos + 1) = Temp3
                Temp4 = Names(Pos)
                Names(Pos) = Names(Pos + 1)
                Names(Pos + 1) = Temp4
            End If
        Next Pos
    Next Pass
    
    ' print the sorted list
    picHighScores.Print "Recent High Scores (highest to lowest):"
    picHighScores.Print
    picHighScores.Print "Rank"; Tab(10); "Player"; Tab(30); "Score"
    picHighScores.Print
    picHighScores.Print "**********************************************************************"
    
    For Increment = 1 To 10
        picHighScores.Print FormatNumber(Increment, 0) & "."; Tab(10); Names(Increment); Tab(30); FormatCurrency(HighScores(Increment), 0)
    Next Increment
End Sub



