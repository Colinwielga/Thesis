VERSION 5.00
Begin VB.Form FrmNFL 
   BackColor       =   &H80000017&
   Caption         =   "Form NFL"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13230
   LinkTopic       =   "Form2"
   ScaleHeight     =   10500
   ScaleWidth      =   13230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   8160
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   3240
      Picture         =   "football.frx":0000
      ScaleHeight     =   1035
      ScaleWidth      =   9675
      TabIndex        =   7
      Top             =   120
      Width           =   9735
   End
   Begin VB.CommandButton cmdsuperbowl 
      Caption         =   "Who will win the Super Bowl XXXVIII?"
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   6960
      Width           =   2775
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   9240
      Width           =   2775
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search For Team"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   5880
      Width           =   2775
   End
   Begin VB.CommandButton cmdsortpower 
      Caption         =   "Sort By Power Ranking"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   2775
   End
   Begin VB.CommandButton cmdsortpercentage 
      Caption         =   "Sort By Percentage"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   2775
   End
   Begin VB.CommandButton cmdSortname 
      Caption         =   "Sort By Name"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   2775
   End
   Begin VB.PictureBox pbxResults 
      Height          =   8775
      Left            =   3240
      ScaleHeight     =   8715
      ScaleWidth      =   9675
      TabIndex        =   0
      Top             =   1320
      Width           =   9735
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Creator: Adam Hanson"
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Welcome to the NFL Search and Sort Engine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "FrmNFL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NFL Search and Sort Engine (NFL Project)
'Form NFL (FrmNFL)
'Author Adam Hanson
'Date Written 10/25/03
'Purpose to be able to search and organize teams by name, win percentage, and my own specail power ranking

Option Explicit
Dim team(1 To 100) As String
Dim wins(1 To 100) As Integer
Dim losses(1 To 100) As Integer
Dim percentage(1 To 100) As Single
Dim pointsfor(1 To 100) As Integer
Dim pointsagainst(1 To 100) As Integer
Dim strpath As String
Dim strfile As String




Private Sub cmdclear_Click()
    pbxResults.Cls
End Sub

Private Sub cmdquit_Click()
    End
End Sub

Private Sub cmdsearch_Click() ' to search by team name
    Dim searchteam As String
    Dim i As Integer
    Dim powerranking(1 To 32) As Double
    pbxResults.Cls
    pbxResults.Print "Team Name", Tab(29); "Wins", "Losses", "Win %", "PF", "PA", "Power Ranking"
    pbxResults.Print " "
    strfile = strpath & "data.txt"
    Open strfile For Input As #1
    For i = 1 To 32
          Input #1, team(i), wins(i), losses(i), percentage(i), pointsfor(i), pointsagainst(i)
        powerranking(i) = (pointsfor(i) - pointsagainst(i)) / pointsagainst(i) 'power ranking is my own special ranking which effectivly ranks teams by points for and points against.  If you look closley the teams with the best ranking are generally the teams with the best win percentage
    Next i
    searchteam = InputBox("Enter a Team Name:")
    i = 1
        Do Until team(i) = searchteam Or i = 33
            i = i + 1
        Loop
    If i = 33 Then
        pbxResults.Print "Team Not Found"
    Else
        pbxResults.Print team(i), wins(i), losses(i), FormatPercent(percentage(i)), pointsfor(i), pointsagainst(i), FormatNumber(powerranking(i))
    End If
    Close #1


End Sub

Private Sub cmdSortname_Click() ' sort by name
Dim N As Integer
Dim temp As String
Dim i As Integer
Dim Pass As Integer
Dim powerranking(1 To 32) As Double
strfile = strpath & "data.txt"
Open strfile For Input As #1
pbxResults.Cls
For i = 1 To 32
    Input #1, team(i), wins(i), losses(i), percentage(i), pointsfor(i), pointsagainst(i)
    powerranking(i) = (pointsfor(i) - pointsagainst(i)) / pointsagainst(i)
Next i
    pbxResults.Print "Team Name", Tab(29); "Wins", "Losses", "Win %", "PF", "PA", "Power Ranking"
    pbxResults.Print " "
    N = 32
    For Pass = 1 To N - 1
    For i = 1 To N - Pass
    If team(i) > team(i + 1) Then
            temp = team(i)
            team(i) = team(i + 1)
            team(i + 1) = temp
            
            temp = wins(i)
            wins(i) = wins(i + 1)
            wins(i + 1) = temp

            temp = losses(i)
            losses(i) = losses(i + 1)
            losses(i + 1) = temp

            temp = percentage(i)
            percentage(i) = percentage(i + 1)
            percentage(i + 1) = temp

            temp = pointsfor(i)
            pointsfor(i) = pointsfor(i + 1)
            pointsfor(i + 1) = temp
            
            temp = pointsagainst(i)
            pointsagainst(i) = pointsagainst(i + 1)
            pointsagainst(i + 1) = temp
            
            temp = powerranking(i)
            powerranking(i) = powerranking(i + 1)
            powerranking(i + 1) = temp
    
            End If
    Next i
    Next Pass
    For i = 1 To N
    pbxResults.Print team(i), wins(i), losses(i), FormatPercent(percentage(i)), pointsfor(i), pointsagainst(i), FormatNumber(powerranking(i))
    Next i
    Close #1
End Sub

Private Sub cmdsortpercentage_Click() ' sort by win perentage
Dim N As Integer
Dim temp As String
Dim i As Integer
Dim Pass As Integer
Dim powerranking(1 To 32) As Double
strfile = strpath & "data.txt"
Open strfile For Input As #1
pbxResults.Cls
For i = 1 To 32
    Input #1, team(i), wins(i), losses(i), percentage(i), pointsfor(i), pointsagainst(i)
    powerranking(i) = (pointsfor(i) - pointsagainst(i)) / pointsagainst(i)
Next i
    pbxResults.Print "Team Name", Tab(29); "Wins", "Losses", "Win %", "PF", "PA", "Power Ranking"
    
    N = 32
    For Pass = 1 To N - 1
    For i = 1 To N - Pass
    If percentage(i) < percentage(i + 1) Then
            temp = percentage(i)
            percentage(i) = percentage(i + 1)
            percentage(i + 1) = temp
            
            temp = wins(i)
            wins(i) = wins(i + 1)
            wins(i + 1) = temp

            temp = losses(i)
            losses(i) = losses(i + 1)
            losses(i + 1) = temp

            temp = team(i)
            team(i) = team(i + 1)
            team(i + 1) = temp

            temp = pointsfor(i)
            pointsfor(i) = pointsfor(i + 1)
            pointsfor(i + 1) = temp
            
            temp = pointsagainst(i)
            pointsagainst(i) = pointsagainst(i + 1)
            pointsagainst(i + 1) = temp

            temp = powerranking(i)
            powerranking(i) = powerranking(i + 1)
            powerranking(i + 1) = temp
            
            End If
    Next i
    Next Pass
    For i = 1 To N
        pbxResults.Print team(i), wins(i), losses(i), FormatPercent(percentage(i)), pointsfor(i), pointsagainst(i), FormatNumber(powerranking(i))
    Next i
    Close #1

End Sub

Private Sub cmdsortpower_Click() ' search by power ranking explained above
Dim N As Integer
Dim temp As String
Dim i As Integer
Dim Pass As Integer
Dim powerranking(1 To 32) As Double
strfile = strpath & "data.txt"
Open strfile For Input As #1
pbxResults.Cls
For i = 1 To 32
    Input #1, team(i), wins(i), losses(i), percentage(i), pointsfor(i), pointsagainst(i)
    powerranking(i) = (pointsfor(i) - pointsagainst(i)) / pointsagainst(i)
Next i
    pbxResults.Print "Team Name", Tab(29); "Wins", "Losses", "Win %", "PF", "PA", "Power Ranking"
    
    N = 32
    For Pass = 1 To N - 1
    For i = 1 To N - Pass
    If powerranking(i) < powerranking(i + 1) Then
            temp = powerranking(i)
            powerranking(i) = powerranking(i + 1)
            powerranking(i + 1) = temp
            
            
            temp = percentage(i)
            percentage(i) = percentage(i + 1)
            percentage(i + 1) = temp
            
            temp = wins(i)
            wins(i) = wins(i + 1)
            wins(i + 1) = temp

            temp = losses(i)
            losses(i) = losses(i + 1)
            losses(i + 1) = temp

            temp = team(i)
            team(i) = team(i + 1)
            team(i + 1) = temp

            temp = pointsfor(i)
            pointsfor(i) = pointsfor(i + 1)
            pointsfor(i + 1) = temp
            
            temp = pointsagainst(i)
            pointsagainst(i) = pointsagainst(i + 1)
            pointsagainst(i + 1) = temp

            End If
    Next i
    Next Pass
    For i = 1 To N
        pbxResults.Print team(i), wins(i), losses(i), FormatPercent(percentage(i)), pointsfor(i), pointsagainst(i), FormatNumber(powerranking(i))
    Next i
    Close #1
End Sub

Private Sub cmdsuperbowl_Click()
     FrmIndex.Show
     FrmNFL.Hide
End Sub

Private Sub Form_Load()
    strpath = "N:\CS130\handin\athanson\"
End Sub
