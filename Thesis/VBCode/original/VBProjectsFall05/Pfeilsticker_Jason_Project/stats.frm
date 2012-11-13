VERSION 5.00
Begin VB.Form PositionPlayers 
   BackColor       =   &H00400000&
   Caption         =   "Form1"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox piclogo 
      Height          =   1575
      Left            =   240
      Picture         =   "stats.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   7680
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdgoto 
      Caption         =   "Go To Pitching Stats"
      Height          =   855
      Left            =   7680
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdBA 
      Caption         =   "Sort by Batting Average"
      Height          =   855
      Left            =   7440
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Player's Stats"
      Height          =   855
      Left            =   6000
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      Height          =   855
      Left            =   4560
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdgetstats 
      Caption         =   "GET STATS FOR 2005 SEASON"
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      Height          =   3975
      Left            =   480
      ScaleHeight     =   3915
      ScaleWidth      =   8715
      TabIndex        =   0
      Top             =   3360
      Width           =   8775
   End
   Begin VB.Label lblMN 
      BackColor       =   &H00400000&
      Caption         =   "Minnesota Twins Batters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   2280
      TabIndex        =   10
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label lbldesigner 
      Caption         =   "Designed:  Jason Pfeilsticker"
      Height          =   255
      Left            =   7200
      TabIndex        =   9
      Top             =   7560
      Width           =   2055
   End
   Begin VB.Label lblName 
      BackColor       =   &H00400000&
      Caption         =   "Enter the name of the player that you would like to find"
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   2880
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
End
Attribute VB_Name = "PositionPlayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form will get stats for Minnesota Twins Players for the 2005 season.
'This for in particular will deal with batting statistics from the season.
'A person will be able to type in a name to look for a player and will be able to sort by highes batting average

Private Sub cmdBA_Click()
'clear anything in the picture box
picResults.Cls
'Sorts by AVG, Batter with highest AVG will be first
'then also have to sort all the other stats so that they go with the correct player

For Pass = 1 To CTR - 1
    For Comp = 1 To CTR - Pass
        If AVG(Comp) < AVG(Comp + 1) Then
            tempAVG = AVG(Comp)
            AVG(Comp) = AVG(Comp + 1)
            AVG(Comp + 1) = tempAVG
            tempBatters = Batters(Comp)
            Batters(Comp) = Batters(Comp + 1)
            Batters(Comp + 1) = tempBatters
            tempGames = Games(Comp)
            Games(Comp) = Games(Comp + 1)
            Games(Comp + 1) = tempGames
            tempAtBats = AtBats(Comp)
            AtBats(Comp) = AtBats(Comp + 1)
            AtBats(Comp + 1) = tempAtBats
            tempHits = Hits(Comp)
            Hits(Comp) = Hits(Comp + 1)
            Hits(Comp + 1) = tempHits
            tempHR = HR(Comp)
            HR(Comp) = HR(Comp + 1)
            HR(Comp + 1) = tempHR
            tempRBI = RBI(Comp)
            RBI(Comp) = RBI(Comp + 1)
            RBI(Comp + 1) = tempRBI
        End If
    Next Comp
Next Pass

'print headers
picResults.Print "Player"; Tab(30); "Games"; Tab(45); "At Bats"; Tab(60); "Hits"; Tab(75); "Home Runs"; Tab(90); "RBI"; Tab(105); "AVG"
'run a loop to print out the names of players and their batting average
For J = 1 To CTR
             picResults.Print Batters(J), Tab(30); Games(J), Tab(45); AtBats(J), Tab(60); Hits(J), Tab(75); HR(J), Tab(90); RBI(J), Tab(105); AVG(J)
Next J
End Sub

Private Sub cmdgetstats_Click()
Open App.Path & "\battingstats.txt" For Input As #1
Open App.Path & "\pitchingstats.txt" For Input As #2
'open files to be put into array
'put stats from the batting file into an array
'CTR gives value to the array positon
CTR = 0
CTR2 = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Batters(CTR), Games(CTR), AtBats(CTR), Hits(CTR), HR(CTR), RBI(CTR), AVG(CTR)
Loop
Close (1)
'Close(1) closes the file
'put stats from the pitching file into an array
'CTR2 gives value to the array position
Do While Not EOF(2)
    CTR2 = CTR2 + 1
    Input #2, Pitchers(CTR2), Wins(CTR2), Losses(CTR2), ERA(CTR2), Saves(CTR2), Innings(CTR2), Strikeouts(CTR2)
Loop
Close (2)
'Close(2) closes the file


End Sub
Private Sub cmdgoto_Click()
'this will go to the pitching stats form
Pitching.Show
PositionPlayers.Hide
End Sub


Private Sub cmdName_Click()
'clears anything in the picture box
picResults.Cls

'this will find the player's stats when a user enters the name in the text box
I = 0
NotFound = True
Player = txtName.Text
'goes through a loop to find the player
Do While NotFound And I < CTR
    I = I + 1
    If Player = Batters(I) Then NotFound = False
Loop
If NotFound Then
        'message box for person not found on the Minnesota Twins
        MsgBox "The name of the person you entered is not on the Minnesota Twins roster", , "Error"
    Else
        picResults.Print "Player"; Tab(15); Tab(30); "Games"; Tab(45); "At Bats"; Tab(60); "Hits"; Tab(75); "Home Runs"; Tab(90); "RBI"; Tab(105); "AVG"
        'print statement for found player
        picResults.Print Batters(I), Tab(15); Tab(30); Games(I), Tab(45); AtBats(I), Tab(60); Hits(I), Tab(75); HR(I), Tab(90); RBI(I), Tab(105); AVG(I)
End If


End Sub

Private Sub cmdQuit_Click()
End
'ends the program
End Sub

