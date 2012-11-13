VERSION 5.00
Begin VB.Form frmBaseball 
   Caption         =   "Baseball Statistics"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   Picture         =   "frmBaseball.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdLookup 
      Caption         =   "Who would you like to lookup?"
      Height          =   615
      Left            =   600
      TabIndex        =   8
      Top             =   7920
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Go back to Main Menu"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   7080
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      Height          =   8295
      Left            =   3720
      ScaleHeight     =   8235
      ScaleWidth      =   5595
      TabIndex        =   6
      Top             =   360
      Width           =   5655
   End
   Begin VB.CommandButton cmdStrikeouts 
      Caption         =   "Which pitcher's have the most strikeouts of all time?"
      Height          =   975
      Left            =   600
      TabIndex        =   5
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton cmdHR 
      Caption         =   "Who has hit the most home runs?"
      Height          =   735
      Left            =   600
      TabIndex        =   4
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdAvg 
      Caption         =   "Who has the highest batting average?"
      Height          =   735
      Left            =   600
      TabIndex        =   3
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdERA 
      Caption         =   "Who has the lowest ERA?"
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdPitcherWins 
      Caption         =   "Which pitcher has the most wins ever?"
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdGames 
      Caption         =   "Who has the most games played?"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmBaseball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'These forms all read in information from a text file and then
'put it into an array and displays it in the picture box
Private Sub cmdAvg_Click()
    picResults.Cls
    Dim CTR As Integer, pass As Integer
    Open App.Path & "\BattingAvg.txt" For Input As #1
    CTR = 0
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, AvgNames(CTR), Average(CTR)
    Loop
    Close #1
    picResults.Print "(note that players in all capitals are currently playing)"
    picResults.Print "Player Names"; Tab(25); "Average"; Tab(40); "Rank"
    picResults.Print "****************************************************************************"
    For pass = 1 To CTR
    '   here we used format number to include the third decimal place for numbers which were only one or two decimal space
        picResults.Print AvgNames(pass); Tab(25); FormatNumber(Average(pass), 3); Tab(40); pass
    Next pass
End Sub

Private Sub cmdERA_Click()
    'This subroutine will read a file into an array and then display it
    picResults.Cls
    Dim CTR As Integer, pass As Integer
    Open App.Path & "\ERA.txt" For Input As #1
    CTR = 0
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, ERANames(CTR), ERA(CTR)
    Loop
    Close #1
    picResults.Print "(note that players in all capitals are currently playing)"
    picResults.Print "Player Names"; Tab(25); "ERA"; Tab(40); "Rank"
    picResults.Print "****************************************************************************"
    For pass = 1 To CTR
    '   here we used format number to include the second decimal place for numbers which were only one decimal space
        picResults.Print ERANames(pass); Tab(25); FormatNumber(ERA(pass), 2); Tab(40); pass
    Next pass
End Sub

Private Sub cmdGames_Click()
    'This subroutine will read a file into an array and then display it

    picResults.Cls
    Dim CTR As Integer, pass As Integer
    Open App.Path & "\GamesPlayed.txt" For Input As #1
    CTR = 0
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, BaseballNames(CTR), GamesPlayed(CTR), Rank(CTR)
    Loop
    Close #1
    picResults.Print "(note that players in all capitals are currently playing)"
    picResults.Print "Player Names"; Tab(25); "Games Played"; Tab(40); "Rank"
    picResults.Print "****************************************************************************"
    For pass = 1 To CTR
        picResults.Print BaseballNames(pass); Tab(25); GamesPlayed(pass); Tab(40); Rank(pass)
    Next pass
    
End Sub

Private Sub cmdHR_Click()
    'This subroutine will read a file into an array and then display it
    picResults.Cls
    Dim CTR As Integer, pass As Integer
    Open App.Path & "\HomeRuns.txt" For Input As #1
    CTR = 0
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, HRNames(CTR), HR(CTR)
    Loop
    Close #1
    picResults.Print "(note that players in all capitals are currently playing)"
    picResults.Print "Player Names"; Tab(25); "Home Runs"; Tab(40); "Rank"
    picResults.Print "****************************************************************************"
    For pass = 1 To CTR
        picResults.Print HRNames(pass); Tab(25); HR(pass); Tab(40); pass
    Next pass
End Sub

Private Sub cmdLookup_Click()
    'This subroutine will allow the user to input a players name from our data file
    'and then displays the players statistics, based on whether they are a pitcher or batter
    Dim found As String, pass As Integer, x As Integer, gate As String
    picResults.Cls
    pass = 0
    found = InputBox("who would you like to search?")
    gate = InputBox("if he is a batter enter 1, if he is a pitcher enter 2")
    For x = 1 To 25
        If (found = PitcherNames(x) Or found = KNames(x) Or found = BaseballNames(x) Or found = HRNames(x) Or found = ERANames(x) Or found = AvgNames(x)) Then
            If (gate = 1) Then
        picResults.Print found; " has "; HR(x); "Home runs and hit for "; Average(x); " average"
            End If
            If (gate = 2) Then
        picResults.Print found; " has "; Strikeouts(x); "strikeouts and has "; PitcherWins(x); " wins"
            End If
        Else
            pass = pass + 1
       End If
    Next x
   If (pass = 25) Then
    MsgBox "invalid entry"
    End If
        
    
End Sub

Private Sub cmdPitcherWins_Click()
    'This subroutine will read a file into an array and then display it
    picResults.Cls
    Dim CTR As Integer, pass As Integer
    Open App.Path & "\PitchWins.txt" For Input As #1
    CTR = 0
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, PitcherNames(CTR), PitcherWins(CTR)
    Loop
    Close #1
    picResults.Print "(note that players in all capitals are currently playing)"
    picResults.Print "Player Names"; Tab(25); "Games Won"; Tab(40); "Rank"
    picResults.Print "****************************************************************************"
    For pass = 1 To CTR
        picResults.Print PitcherNames(pass); Tab(25); PitcherWins(pass); Tab(40); pass
    Next pass
    
End Sub

Private Sub cmdReturn_Click()
    'This subroutine will return you to the Main Menu
    frmBaseball.Hide
    frmHome.Show
End Sub

Private Sub cmdStrikeouts_Click()
    'This subroutine will read a file into an array and then display it
    picResults.Cls
    Dim CTR As Integer, pass As Integer
    Open App.Path & "\Strikeouts.txt" For Input As #1
    CTR = 0
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, KNames(CTR), Strikeouts(CTR)
    Loop
    Close #1
    picResults.Print "(note that players in all capitals are currently playing)"
    picResults.Print "Player Names"; Tab(25); "Strikeouts"; Tab(40); "Rank"
    picResults.Print "****************************************************************************"
    For pass = 1 To CTR
        picResults.Print KNames(pass); Tab(25); Strikeouts(pass); Tab(40); pass
    Next pass
End Sub
