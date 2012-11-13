VERSION 5.00
Begin VB.Form Stats 
   BackColor       =   &H0000C000&
   Caption         =   "Form1"
   ClientHeight    =   5385
   ClientLeft      =   1110
   ClientTop       =   5565
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   7995
   Begin VB.CommandButton SortbyAVG 
      Caption         =   "Sort the team by batting Average"
      Height          =   615
      Left            =   2160
      TabIndex        =   7
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton battingAvg 
      Caption         =   "Calculates the Batting Average"
      Height          =   615
      Left            =   2160
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton BestAVG 
      Caption         =   "Find the Best Average on the Team"
      Height          =   615
      Left            =   2160
      TabIndex        =   5
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton findXsBA 
      Caption         =   "Find A Person's Batting Average"
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton TeamStats 
      Caption         =   "Print Team Stats"
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton GetInfo 
      BackColor       =   &H00000000&
      Caption         =   "Load the Array"
      Height          =   615
      Left            =   2160
      MaskColor       =   &H00808080&
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0080FF80&
      Height          =   4455
      Left            =   3720
      ScaleHeight     =   4395
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Stats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BaseballStatsProject : (BaseballStats1.vbp)
'Form Name: Stats (BaseballStats.frm)
'Author: Peter Scheffler
'Written 3/14/04
'Purpose of project: To find the batting averages of
                    'people on St.Patrick Irish DRS team
                    'print the hitting stats in a table
                    'organize the stats by batting average
                    'find and name the player with the best
                    'average on the team
Option Explicit
Dim Names(1 To 20) As String
Dim AB(1 To 20) As Integer
Dim Hits(1 To 20) As Integer
Dim AVG(1 To 20) As Single






Private Sub battingAvg_Click()
'Calculates the batting averages
Dim J As Integer
Dim CTR As Integer
picResults.Cls
For CTR = 1 To 20
    'Calculates the batting average(DOES NOT Work)
    AVG(CTR) = Hits(CTR) / AB(CTR)
    picResults.Print Names(CTR); Tab(20); AB(CTR), Hits(CTR), FormatNumber(AVG(CTR), 3)
Next CTR
battingAvg.Visible = False
TeamStats.Visible = True
findXsBA.Visible = True
BestAVG.Visible = True
SortbyAVG.Visible = True
End Sub



Private Sub Form_Load()
battingAvg.Visible = False
TeamStats.Visible = False
findXsBA.Visible = False
BestAVG.Visible = False
SortbyAVG.Visible = False
End Sub

Private Sub TeamStats_Click()
'Computes the player's Batting Average
'Prints out the All player's Names,
'At Bats, Hits, and Batting Average
Dim J As Integer
Dim CTR As Integer
picResults.Cls
For CTR = 1 To 20
    'Calculates the batting average(DOES NOT Work)
    AVG(CTR) = Hits(CTR) / AB(CTR)
    picResults.Print FormatNumber(AVG(CTR), 3)
Next CTR
End Sub



Private Sub BestAVG_Click()
'Finds the best Average on the team
picResults.Cls
Dim BestAverage As Single
Dim BestHitter As String
Dim CTR As Integer
Dim I As Integer
BestAverage = AVG(1)
BestHitter = Names(1)
CTR = 0
picResults.Cls
For I = 2 To 20
    If AVG(I) > BestAverage Then
        BestAverage = AVG(I)
        BestHitter = Names(I)
    End If
Next I
For I = 1 To 20
    picResults.Print I; "."; Names(I); Tab(30); FormatNumber((AVG(I)), 3)
Next I
picResults.Print BestHitter; " has the best average of "; FormatNumber((BestAverage), 3)
BestAVG.Visible = False
End Sub

Private Sub findXsBA_Click()
'Finds out if a user entered name is on the team
Dim Position As Integer
Dim Found As Boolean
Dim Batter As String
Found = False
Position = 0
picResults.Cls
Batter = InputBox("Enter a Player's Name", "Name")
Do While (Not Found) And (Position < 20)
    Position = Position + 1
    If Names(Position) = Batter Then
        picResults.Print "The Batting Average for"; Names(Position); " is"; AVG(Position)
        Found = True
    End If
Loop
If Found = False Then
    MsgBox "Sorry but the Batter entered is not on the Team", , "Error"
End If
End Sub

Private Sub GetInfo_Click()
'Loads the Batting information from notepad
'into an Array
Dim CTR As Integer
CTR = 0
Open "N:\CS130\handin\Scheffler,Peter\Stats.txt" For Input As #1

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Names(CTR), AB(CTR), Hits(CTR)
Loop
Close #1
battingAvg.Visible = True


End Sub

Private Sub Quit_Click()
End
End Sub

Private Sub SortbyAVG_Click()
Dim Pass As Integer
Dim Temp As Single
Dim CTR As Integer
Dim TempNames As String
picResults.Cls
Pass = 0
CTR = 0
'Sorts the list from greatest to least AVG
For Pass = 1 To 19
    For CTR = 1 To (20 - Pass)
        If AVG(CTR) < AVG(CTR + 1) Then
            Temp = AVG(CTR)
            AVG(CTR) = AVG(CTR + 1)
            AVG(CTR + 1) = Temp
                TempNames = Names(CTR)
                Names(CTR) = Names(CTR + 1)
                Names(CTR + 1) = TempNames
        End If
    Next CTR
Next Pass
'Output the Averages in order from greatest to least
For CTR = 1 To 20
    picResults.Print Names(CTR); Tab(20); FormatNumber((AVG(CTR)), 3)
Next CTR
SortbyAVG.Visible = False
End Sub


