VERSION 5.00
Begin VB.Form frmTeamStats
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   11835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17280
   LinkTopic       =   "Form1"
   Picture         =   "frmTeamStats.frx":0000
   ScaleHeight     =   11835
   ScaleWidth      =   17280
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResultsBA
      BackColor       =   &H0080FFFF&
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
      Left            =   8640
      ScaleHeight     =   1155
      ScaleWidth      =   6435
      TabIndex        =   2
      Top             =   6240
      Width           =   6495
   End
   Begin VB.CommandButton cmdTeamBA
      BackColor       =   &H0000FFFF&
      Caption         =   "Calculate Team Batting Average"
      BeginProperty Font
         Name            =   "Castellar"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   3375
   End
   Begin VB.CommandButton cmdReturn
      BackColor       =   &H0000FFFF&
      Caption         =   "Return to Main Page"
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
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10440
      Width           =   2535
   End
   Begin VB.Image Image1
      Height          =   1470
      Left            =   15480
      Picture         =   "frmTeamStats.frx":2A9802
      Top             =   10320
      Width           =   1680
   End
End
Attribute VB_Name = "frmTeamStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Team stats page - take team averages of given stats

'Returns to main page
Private Sub cmdReturn_Click()
frmDeffensivePage.Hide
frmPitchingPage.Hide
frmOffensivePage.Hide
frmAHomePage.Show
End Sub

Private Sub cmdTeamBA_Click()
Open App.Path & "\OffensiveStats.txt" For Input As #1
'intializing counter
Counter = 0
'opening file
Dim TeamTotal As Single
Dim Average(1 To 20) As Single
Dim TeamAverage As Single
Dim I As Integer
 Do While Not EOF(1)
    Counter = Counter + 1
       Input #1, PlayerNumber(Counter), PlayerName(Counter), PA(Counter), AB(Counter), Hits(Counter), Doubles(Counter), Triples(Counter), HR(Counter), BB(Counter), KO(Counter), BattingAverage(Counter)
        Average(Counter) = Hits(Counter) / AB(Counter)
Loop

'calculate team batting average
For I = 1 To Counter
TeamTotal = TeamTotal + Average(I)

Next I

TeamAverage = TeamTotal / Counter

'print results
picResultsBA.Print "The team batting average is "; FormatNumber(TeamAverage, 3); "."

End Sub

Private Sub cmdTeamFP_Click()
'intializing counter
Counter = 0
'opening file
Open App.Path & "\DeffensiveStats.txt" For Input As #1
 Do While Not EOF(1)
    Counter = Counter + 1
       Input #1, PutOuts(Counter), Assits(Counter), Errors(Counter)
   Loop
'define variables
Dim FPA As Integer
FPA = PutOuts(Counter) + Assits(Counter)
Dim FPB As Integer
FPB = (PutOuts(Counter) + Assits(Counter) + Errors(Counter))
Dim FieldingPercentage As Single
FieldingPercentage(Counter) = FPA / FPB
TeamFieldingPercentage = FieldingPercentage(Counter)
Dim TeamAverage As Single
TeamAverage = TeamFieldingPercentage / Counter
'print results
Close #1
picResultsBA.Print "The team batting average is "; FormatPercent(TeamAverage); "."
End Sub

Private Sub cmdTeamSp_Click()
'intializing counter
Counter = 0
'opening file
Open App.Path & "\OffensiveStats.txt" For Input As #1
Dim SluggingPercentage As Single
Dim TeamAverage As Single
Dim I As Integer
Dim TeamTotal As Single
Dim Average(1 To 20) As Integer
 Do While Not EOF(1)
    Counter = Counter + 1
       Input #1, PlayerNumber(Counter), PlayerName(Counter), PA(Counter), AB(Counter), Hits(Counter), Doubles(Counter), Triples(Counter), HR(Counter), BB(Counter), KO(Counter)
    Singles(Counter) = Hits(Counter) - (Doubles(Counter) - Triples(Counter) - HR(Counter))
    SP(Counter) = (Singles(Counter) + 2 * Doubles(Counter) + 3 * Triples(Counter) + 4 * HR(Counter)) / AB(Counter)
    Average(Counter) = SP(Counter) / Counter
    Loop
For I = 1 To Counter
TeamTotal = TeamTotal + Average(I)
Next I

TeamAverage = Average(Counter) / Counter

'print results
picResultsBA.Print "The team slugging percentage is "; FormatNumber(TeamAverage, 3); "."
End Sub

Private Sub Image1_Click()
MsgBox "DLS Softball!"
End Sub
