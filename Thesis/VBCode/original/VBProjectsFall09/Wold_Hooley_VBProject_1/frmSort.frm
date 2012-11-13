VERSION 5.00
Begin VB.Form frmSort 
   BackColor       =   &H80000012&
   Caption         =   "Form1"
   ClientHeight    =   11160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   Picture         =   "frmSort.frx":0000
   ScaleHeight     =   11160
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   7800
      TabIndex        =   3
      Top             =   9840
      Width           =   1935
   End
   Begin VB.CommandButton cmdMainMenu 
      Height          =   1815
      Left            =   360
      Picture         =   "frmSort.frx":7C3C2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9120
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000015&
      FillColor       =   &H0000FFFF&
      ForeColor       =   &H0000FFFF&
      Height          =   7455
      Left            =   4200
      ScaleHeight     =   7395
      ScaleWidth      =   4755
      TabIndex        =   1
      Top             =   1560
      Width           =   4815
   End
   Begin VB.CommandButton cmdSortGoals 
      Height          =   855
      Left            =   600
      Picture         =   "frmSort.frx":87834
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   9375
   End
End
Attribute VB_Name = "frmSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Team(1 To 150) As String, Goals(1 To 150) As Integer, Player(1 To 150) As String, tempGoals As Integer, tempPlayer As String, tempTeam As String, j As Integer
Dim pass As Integer, pos As Integer
'The Purpose of this form is to read a text file with all the teams' players, number of goals and which team they play for and sort them according to most goals _
to fewest then print the top 15.

Private Sub cmdMainMenu_Click()
    frmSort.Hide
    frmMainMenu.Show
End Sub

Private Sub cmdSortGoals_Click()
'This will open the text file and put the information into the appropriate array.
Open App.Path & "\GoalSortFile.txt" For Input As #1
    Ctr = 0
'Print the header
picResults.Print "Player"; Tab(15); "Number of Goals"; Tab(30); "Team"
picResults.Print "***********************************************"

    Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Player(Ctr), Goals(Ctr), Team(Ctr)
        
    Loop
    'Sort from most goals to fewest.
    For pass = 1 To Ctr - 1
        For pos = 1 To Ctr - pass
            If Goals(pos) < Goals(pos + 1) Then
                tempGoals = Goals(pos)
                Goals(pos) = Goals(pos + 1)
                Goals(pos + 1) = tempGoals
                tempPlayer = Player(pos)
                Player(pos) = Player(pos + 1)
                Player(pos + 1) = tempPlayer
                tempTeam = Team(pos)
                Team(pos) = Team(pos + 1)
                Team(pos + 1) = tempTeam
             End If
        Next pos
    Next pass
        
        
    For j = 1 To 15     'This will only print the top 15 goal scorers
        
        picResults.Print Player(j); Tab(20); Goals(j); Tab(30); Team(j)
    
    Next j
 
        
        
  Close #1
        
   

End Sub



Private Sub cmdQuit_Click()

End

End Sub
