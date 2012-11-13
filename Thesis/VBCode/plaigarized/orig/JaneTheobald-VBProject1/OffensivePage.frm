VERSION 5.00
Begin VB.Form frmOffensivePage 
   BackColor       =   &H00000000&
   Caption         =   "Form2"
   ClientHeight    =   10395
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15330
   LinkTopic       =   "Form2"
   Picture         =   "OffensivePage.frx":0000
   ScaleHeight     =   10395
   ScaleWidth      =   15330
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Read 2007 Stats"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox txtDirections 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "OffensivePage.frx":2A9EA6
      Top             =   600
      Width           =   4935
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H0000FFFF&
      Caption         =   "Clear!"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8880
      Width           =   1695
   End
   Begin VB.CommandButton CmdReturn 
      BackColor       =   &H0000FFFF&
      Caption         =   "Return to Main Page"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8880
      Width           =   2055
   End
   Begin VB.CommandButton cmdCalcKO 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sort By Strike Outs"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   3015
   End
   Begin VB.CommandButton cmdcalcWalks 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sort By Number of Walks"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   1470
      Left            =   13440
      Picture         =   "OffensivePage.frx":2A9EF6
      Top             =   8880
      Width           =   1680
   End
End
Attribute VB_Name = "frmOffensivePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Offensive stats page designed to read in a data file with players name number and stats
'With this data allow user to sort players in order of most desireable stats
'Each buttons Results will print to their own data file

Private Sub Check1_Click()
'Intialize Counter
Counter = 0
'Read in file with offensive stats for the year
Open App.Path & "\OffensiveStats.txt" For Input As #1
Do While Not EOF(1)
    Counter = Counter + 1
       Input #1, PlayerNumber(Counter), PlayerName(Counter), PA(Counter), AB(Counter), Hits(Counter), Doubles(Counter), Triples(Counter), HR(Counter), BB(Counter), KO(Counter), BattingAverage(Counter)
Loop
'Tell user file was read
MsgBox "Offensive Stats have been read successfully."
Close #1
End Sub

Private Sub cmdCalcBA_Click()
'sort player by batting average highest to lowest
'declaring variables
Dim Pass As Integer, Pos As Integer, B As Integer
Dim TempBA As Single
Dim TempNumber As Integer
Dim TempName As String

'intialize counter
Counter = Counter + 1
B = B + 1

'open document to print to
Open App.Path & "\PrintBattingAverage.txt" For Output As #1

'bubble sort
For Pass = 1 To Counter - 1
For Pos = 1 To Counter - Pass
        If BattingAverage(Pos) < BattingAverage(Pos + 1) Then
        TempBA = TempBA(Pos)
        TempBA(Pos) = TempBA(Pos + 1)
        TempBA(Pos + 1) = TempBA
        TempNumber = PlayerNumber(Pos)
        PlayerNumber(Pos) = PlayerNumber(Pos + 1)
        PlayerNumber(Pos + 1) = TempNumber
        TempName = PlayerName(Pos)
        PlayerName(Pos) = PlayerName(Pos + 1)
        PlayerName(Pos + 1) = TempName
        End If
    Next Pos
Next Pass

'print results
For B = 1 To 10
    Write #1, PlayerNumber(B), PlayerName(B), BattingAverage(B)
Next B
'tell user at bats have been sorted successfully
MsgBox "Batting Averages have been sorted."
Close #1
End Sub

Private Sub cmdCalcKO_Click()
'sort players by strikeouts lowest to highest
'declaring variables
Dim Pass As Integer, Pos As Integer, K As Integer
Dim TempKO As Integer
Dim TempNumber As Integer
Dim TempName As String
'intialize counter
Counter = Counter + 1
K = K + 1
'print results to output file
Open App.Path & "\PrintStrikeOuts.txt" For Output As #1

'bubble sort
For Pass = 1 To Counter - 1
For Pos = 1 To Counter - Pass
    If KO(Pos) > KO(Pos + 1) Then
        TempKO = KO(Pos)
        KO(Pos) = KO(Pos + 1)
        KO(Pos + 1) = TempKO
        TempNumber = PlayerNumber(Pos)
        PlayerNumber(Pos) = PlayerNumber(Pos + 1)
        PlayerNumber(Pos + 1) = TempNumber
        TempName = PlayerName(Pos)
        PlayerName(Pos) = PlayerName(Pos + 1)
        PlayerName(Pos + 1) = TempName
        End If
    Next Pos
Next Pass

'print results
For K = 2 To 11
    Write #1, PlayerNumber(K), PlayerName(K), KO(K)
Next K
Close #1
'tell user data has been read
MsgBox "Strike Outs have been sorted."
End Sub

Private Sub cmdCalcSP_Click()
'sort players by slugging percentage highest to lowest
'Declaring variables
Dim Pass As Integer, Pos As Integer, S As Integer
Dim TempSP As Integer, TempNumber As Integer
Dim TempName As String

'initialize counter
Counter = Counter + 1
S = S + 1

'print results to output file
Open App.Path & "\PrintSluggingPercentage.txt" For Output As #1
'defining variables
'loops through players
For Counter = 1 To 10
    Singles(Counter) = Hits(Counter) - (Doubles(Counter) - Triples(Counter) - HR(Counter))
    SP(Counter) = Singles(Counter) + 2 * Doubles(Counter) + 3 * Triples(Counter) + 4 * HR(Counter) / AB(Counter)
Next Counter

'bubble sort
For Pass = 1 To Counter - 1
For Pos = 1 To Counter - Pass
        If SP(Pos) < SP(Pos + 1) Then
        TempSP = SP(Pos)
        SP(Pos) = SP(Pos + 1)
        SP(Pos + 1) = TempSP
        TempNumber = PlayerNumber(Pos)
        PlayerNumber(Pos) = PlayerNumber(Pos + 1)
        PlayerNumber(Pos + 1) = TempNumber
        TempName = PlayerName(Pos)
        PlayerName(Pos) = PlayerName(Pos + 1)
        PlayerName(Pos + 1) = TempName
        End If
    Next Pos
Next Pass
'print results
For S = 1 To 11
    Write #1, PlayerNumber(Counter), PlayerName(Counter), SP(Counter)
Next S
'tell user data has been read
MsgBox "Slugging percentages have been sorted."
Close #1
End Sub

Private Sub cmdcalcWalks_Click()
'sort by number of walks highest to lowest
'declaring variables
Dim Pass As Integer, Pos As Integer, W As Integer
Dim TempNumber As Integer, TempWalk As Integer
Dim TempName As String
'intilize counter
Counter = Counter + 1
W = W + 1
'print results to output file
Open App.Path & "\PrintWalks.txt" For Output As #1

'bubble sort
For Pass = 1 To Counter - 1
For Pos = 1 To Counter - Pass
        If BB(Pos) < BB(Pos + 1) Then
        TempWalk = BB(Pos)
        BB(Pos) = BB(Pos + 1)
        BB(Pos + 1) = TempWalk
        TempNumber = PlayerNumber(Pos)
        PlayerNumber(Pos) = PlayerNumber(Pos + 1)
        PlayerNumber(Pos + 1) = TempNumber
        TempName = PlayerName(Pos)
        PlayerName(Pos) = PlayerName(Pos + 1)
        PlayerName(Pos + 1) = TempName
        End If
    Next Pos
Next Pass

'print results
For W = 1 To 10
    Write #1, PlayerNumber(W), PlayerName(W), BB(W)
Next
Close #1
MsgBox "Walks have been sorted."

End Sub

Private Sub cmdClear_Click()
'clear picture box
picOResults.Cls
End Sub


Private Sub cmdReturn_Click()
'Returns to main page
frmAHomePage.Show
frmDeffensivePage.Hide
frmPitchingPage.Hide
frmOffensivePage.Hide
End Sub


Private Sub Image1_Click()
MsgBox "Get a hit!"
End Sub
