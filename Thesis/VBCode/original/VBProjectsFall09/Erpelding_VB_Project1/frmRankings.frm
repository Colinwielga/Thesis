VERSION 5.00
Begin VB.Form frmRankings 
   BackColor       =   &H80000007&
   Caption         =   "Conference Rankings"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10605
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWins 
      BackColor       =   &H00000080&
      Caption         =   "Display Team's Wins and Losses"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   3195
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00000080&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7200
      Width           =   3255
   End
   Begin VB.TextBox txtTitle 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Berlin Sans FB Demi"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1680
      TabIndex        =   4
      Text            =   "NIVC Rankings"
      Top             =   480
      Width           =   6855
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00000080&
      Caption         =   "Quit All Forms"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8760
      Width           =   3255
   End
   Begin VB.CommandButton cmdDisplay1 
      BackColor       =   &H00000080&
      Caption         =   "Display Records in Order from Best to Worst"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   3255
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H00000080&
      Caption         =   "Read Conference Standings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   3255
   End
   Begin VB.PictureBox picResults 
      Height          =   8175
      Left            =   4920
      ScaleHeight     =   8115
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   2040
      Width           =   4935
   End
End
Attribute VB_Name = "frmRankings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ctr As Integer, Teams(1 To 13) As String, Wins(1 To 13) As Integer, Losses(1 To 13) As Integer, WinPercentage(1 To 13) As Single


'clear the picbox
Private Sub cmdClear_Click()
picResults.Cls
End Sub

'This button will display the names of the schools in the conference along with their wins, losses, and winning percentage from last year
Private Sub cmdDisplay_Click()
'Declare variables and arrays
Dim J As Integer, SchoolNames(1 To 13) As String, Wins(1 To 13) As Integer, Losses(1 To 13) As Integer, WinPercentage(1 To 13) As Single

'Print headers
picResults.Print "School Name", "# of Wins", "# of Losses", "Winning Percentage"
picResults.Print "******************************************************************************"

Do While Not EOF(2)
    Ctr = Ctr + 1
Loop

picResults.Print SchoolNames(Ctr), Wins(Ctr), Losses(Ctr), WinPercentage(Ctr)

End Sub

'This button will display the teams in order of best to worst based on their winning percentage
Private Sub cmdDisplay1_Click()
'Declare variables
Dim Pass As Integer, Pos As Integer, TempWinPercentage As Integer, I As Integer, TempTeams As String, TempWins As Integer, TempLosses As Integer

For Pass = 1 To Ctr - 1                                 'using the bubble sort to put teams in order based on winning percentage
    For Pos = 1 To Ctr - Pass
        If WinPercentage(Pos) < WinPercentage(Pos + 1) Then
            TempWinPercentage = WinPercentage(Pos)
            WinPercentage(Pos) = WinPercentage(Pos + 1)
            WinPercentage(Pos + 1) = TempWinPercentage
            TempTeams = Teams(Pos)
            Teams(Pos) = Teams(Pos + 1)
            Teams(Pos + 1) = TempTeams
            TempWins = Wins(Pos)
            Wins(Pos) = Wins(Pos + 1)
            Wins(Pos + 1) = TempWins
            TempLosses = Losses(Pos)
            Losses(Pos) = Losses(Pos + 1)
            Losses(Pos + 1) = TempLosses
        End If
    Next Pos
Next Pass
'print headers
picResults.Print "Teams"; Tab(20); "Wins"; Tab(30); "Losses"; Tab(40); "WinPercentage"
picResults.Print "*************************************************************************"

'Print the sorted list
For I = 1 To Ctr
    picResults.Print Teams(I); Tab(20); Wins(I); Tab(30); Losses(I); Tab(40); WinPercentage(I)
Next I

        
End Sub

'End the Program
Private Sub cmdQuit_Click()
    End
End Sub
'Read the file and put into arrays
Private Sub cmdRead_Click()

'Open file
Open App.Path & "\Rankings.txt" For Input As #2

'Read through the file and put into correct arrays
Do While Not EOF(2)
    Ctr = Ctr + 1
    Input #2, Teams(Ctr), Wins(Ctr), Losses(Ctr), WinPercentage(Ctr)
Loop

'Display a message box saying that the file has been read into arrays
MsgBox ("The text file has been read into separate arrays")


cmdDisplay1.Enabled = True
cmdWins.Enabled = True
End Sub

Private Sub cmdWins_Click()
Dim I As Integer
'Print headers
picResults.Print "Teams"; Tab(20); "Wins"; Tab(30); "Losses"
picResults.Print "********************************************"

For I = 1 To Ctr
    picResults.Print Teams(I); Tab(20); Wins(I); Tab(30); Losses(I)
Next I

End Sub
