VERSION 5.00
Begin VB.Form frmAllTeams 
   BackColor       =   &H00FFC0FF&
   Caption         =   "All Teams"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10815
   FillColor       =   &H008080FF&
   ForeColor       =   &H00C0C0FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   10815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H0080FFFF&
      Caption         =   "Search"
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox txtRank 
      Height          =   735
      Left            =   600
      TabIndex        =   7
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080FFFF&
      Caption         =   "Back"
      Height          =   615
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      Height          =   6135
      Left            =   2760
      ScaleHeight     =   6075
      ScaleWidth      =   4635
      TabIndex        =   4
      Top             =   120
      Width           =   4695
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0080FFFF&
      Caption         =   "Exit"
      Height          =   615
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H0080FFFF&
      Caption         =   "Show Total Stats"
      Height          =   735
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdAway 
      BackColor       =   &H0080FFFF&
      Caption         =   "Show Away Stats"
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H0080FFFF&
      Caption         =   "Show Home Stats"
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label lblbase 
      Alignment       =   2  'Center
      Caption         =   "Baseball is Sweet!  It is my second favorite sport!"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7680
      TabIndex        =   10
      Top             =   360
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   4365
      Left            =   7560
      Picture         =   "frmAllTeams.frx":0000
      Top             =   1920
      Width           =   3180
   End
   Begin VB.Label lblrank 
      Alignment       =   2  'Center
      Caption         =   "Find Home Attendance Percentage Greater than Number"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label lblTylerFlory 
      Caption         =   "Tyler Flory"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmAllTeams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'print the team's away stats
Private Sub cmdAway_Click()
    Dim pos As Integer
       picResults.Cls
      picResults.Print "Team", "AwayGames", "AwayAvg", "AwayPercent"
    For pos = 1 To 12
        picResults.Print Team(pos), AwayGames(pos), AwayAverage(pos), AwayPercent(pos)
    Next pos
End Sub
'go back to the start form
Private Sub cmdBack_Click()
    frmAllTeams.Visible = False
    frmStart.Visible = True
End Sub
'quit the program
Private Sub cmdExit_Click()
    End
End Sub
'print the team's home stats
Private Sub cmdHome_Click()
    Dim pos As Integer
       picResults.Cls
    picResults.Print "Team", "HomeGames", "HomeTotal", "HomeAverage"
    For pos = 1 To 12
        picResults.Print Team(pos), HomeGames(pos), HomeTotal(pos), HomeAverage(pos)
    Next pos
End Sub
'to search for teams by ranking their home attendance percentage
Private Sub cmdSearch_Click()
Dim temp As String
Dim inside As Single
Dim i As Single
Dim per As Single
i = 0
Dim J As Single
temp = txtRank.Text
   picResults.Cls
   Dim HomePer(1 To 12) As Single
   Dim Team1(1 To 12) As String
For pos = 1 To 12
    If HomePercent(pos) >= temp Then
        i = i + 1
        HomePer(i) = HomePercent(pos)
        Team1(i) = Team(pos)
    End If
    Next pos
 
     For Pass = 1 To i - 1
        For J = 1 To i - Pass
            If HomePer(J) < HomePer(J + 1) Then
                per = HomePer(J)
                HomePer(J) = HomePer(J + 1)
                HomePer(J + 1) = per
                temp = Team1(J)
                Team1(J) = Team1(J + 1)
                Team1(J + 1) = temp
            End If
        Next J
    Next Pass
    For pos = 1 To i
        picResults.Print Team1(pos), FormatPercent(HomePer(pos) / 100)
    Next pos

    




End Sub
'print the total stats for the teams
Private Sub cmdTotal_Click()
    Dim pos As Integer
       picResults.Cls
               picResults.Print "Team", "TotalGames", "TotalAverage", "TotalPercent"
    For pos = 1 To 12
        picResults.Print Team(pos), TotalGames(pos), TotalAverage(pos), TotalPercent(pos)
    Next pos
End Sub

