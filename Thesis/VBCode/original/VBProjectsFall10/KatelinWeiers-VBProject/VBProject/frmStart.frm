VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00400000&
   Caption         =   "Start Form"
   ClientHeight    =   10890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18525
   FillColor       =   &H00800000&
   ForeColor       =   &H00800000&
   LinkTopic       =   "Start Form"
   ScaleHeight     =   10890
   ScaleWidth      =   18525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDataPlayers 
      Caption         =   "Input Current Player Data"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3120
      TabIndex        =   6
      Top             =   3120
      Width           =   3615
   End
   Begin VB.CommandButton cmdVote 
      Caption         =   "Click to Vote for your Favorite Current Player"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      TabIndex        =   5
      Top             =   7080
      Width           =   7455
   End
   Begin VB.CommandButton cmdFormTeam 
      Caption         =   "View Form with Team Record Information"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11880
      TabIndex        =   4
      Top             =   4800
      Width           =   3615
   End
   Begin VB.CommandButton cmdDataTeam 
      Caption         =   "Input Team Record Data"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   11880
      TabIndex        =   3
      Top             =   3120
      Width           =   3615
   End
   Begin VB.CommandButton cmdFormPlayers 
      Caption         =   "View Form with Player Information"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   2
      Top             =   4800
      Width           =   3615
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   15240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   9600
      Width           =   2055
   End
   Begin VB.Image imageTC 
      Height          =   3300
      Left            =   7560
      Picture         =   "frmStart.frx":0000
      Top             =   2520
      Width           =   3435
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00400000&
      Caption         =   "Minnesota Twins: Team Statistics"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   17535
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Initial form displayed in the project

Private Sub cmdDataPlayers_Click() 'Store player data file into an array
'Open file for input
Open App.Path & "\Players.txt" For Input As #1

'initialize variables and use loop to read the file
Ctr = 0
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, PlayerNumber(Ctr), FirstName(Ctr), LastName(Ctr), Position(Ctr), PlayerBattingAvg(Ctr), Birthdate(Ctr)
Loop

MsgBox "The data for each player has been stored." 'notify user the data is now in an array
Close #1

'Enable user to click the button to transfer them to the new form
cmdDataPlayers.Enabled = False
cmdFormPlayers.Enabled = True

End Sub

Private Sub cmdDataTeam_Click() 'input team statistics from file


'open the data file
Open App.Path & "\TeamRecord.txt" For Input As #2

'initialize variables and use loop to read file
CtrTeam = 0
Do Until EOF(2)
    CtrTeam = CtrTeam + 1
    Input #2, Season(CtrTeam), Wins(CtrTeam), Losses(CtrTeam), Attendance(CtrTeam), Champions(CtrTeam)
Loop

MsgBox "The historical team data has been stored." 'notify user that the file has been read into an array
Close #2

'Enable the use to click the button to transfer them to the new form
cmdDataTeam.Enabled = False
cmdFormTeam.Enabled = True

End Sub

Private Sub cmdFormPlayers_Click() 'Show form with player stats
    frmPlayerInfo.Show
    frmStart.Hide
End Sub

Private Sub cmdFormTeam_Click() 'Show form with team stats
    frmTeamStats.Show
    frmStart.Hide
End Sub

Private Sub cmdQuit_Click() 'Exit the project
    End
End Sub

Private Sub cmdVote_Click() 'Show form allowing user to vote for their favorite player
    frmVote.Show
    frmStart.Hide
End Sub
