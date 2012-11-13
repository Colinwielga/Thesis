VERSION 5.00
Begin VB.Form frmrecords 
   BackColor       =   &H00800000&
   Caption         =   "Garrett Sohn"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdtest 
      Caption         =   "Take the Test"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7200
      TabIndex        =   8
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdmain 
      BackColor       =   &H0000FFFF&
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   240
      MaskColor       =   &H0000FFFF&
      TabIndex        =   6
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmddisplay2 
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6840
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmddisplay 
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox playertxt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Text            =   "Players"
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox teamtxt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Text            =   "Teams"
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox picplayer 
      Height          =   3255
      Left            =   4440
      ScaleHeight     =   3195
      ScaleWidth      =   2235
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.PictureBox picteam 
      Height          =   2055
      Left            =   1560
      ScaleHeight     =   1995
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblclick 
      BackColor       =   &H00800000&
      Caption         =   "Click the lists for more information"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   7
      Top             =   2880
      Width           =   2535
   End
End
Attribute VB_Name = "frmrecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'March Madness (madness.vbp)
'Records form (records.frm)
'Garrett Sohn
'March 24, 2006
'This form will list teams and players in separate picture boxes and allows the user to input either a team or a player. When they insert a team or a player additional information is given.
Option Explicit
Dim UCLA As String
Dim N As String
Private Sub cmddisplay_Click()
Dim teams(1 To 10) As String, Pos As Integer, N As Integer
Open App.Path & "\teams.txt" For Input As #1
picplayer.Cls
picteam.Cls
    Pos = 0
    Do While Not EOF(1)
        Pos = Pos + 1
        Input #1, teams(Pos)
    Loop
    For N = 1 To Pos
        picteam.Print teams(N)
    Next N
    Close #1
End Sub

Private Sub cmddisplay2_Click()
Dim players(1 To 15) As String, Pos As Integer, N As Integer
Open App.Path & "\players.txt" For Input As #1
picteam.Cls
picplayer.Cls
    Pos = 0
    Do While Not EOF(1)
        Pos = Pos + 1
        Input #1, players(Pos)
        'sorting through players
    Loop
    For N = 1 To Pos
        picplayer.Print players(N)
    Next N
    Close #1
End Sub

Private Sub cmdmain_Click()
    frmrecords.Hide
    frmmadness.Show
End Sub

Private Sub cmdtest_Click()
    frmrecords.Hide
    frmtest.Show
End Sub


Private Sub picplayer_Click()
    N = InputBox("Please input a player name, you must spell it correctly and with the team name in brackets", "Player")
    If N = "Bill Bradley (Princeton)" Then
        MsgBox "Bill Bradley has Final Four records of 58 points and 22 field goals in a game against Wichita State on March 20, 1965"
    End If
    If N = "Carmelo Anthony (Syracuse)" Then
        MsgBox "Carmelo Anthony holds the single game highest scoring record by a freshman, scoring 33 pionts against Texas on April 5, 2003"
    End If
    If N = "Lennie Rosenbluth (North Carolina)" Then
        MsgBox "Lennie Rosenbluth has the highest attempted field goals in the Final Four with 42.  He played Michigan State on March 22, 1957"
    End If
    If N = "Freddie Banks (UNLV)" Then
        MsgBox " Freddia Banks made 10 three-pointers against Indiana on March 28, 1987"
    End If
    If N = "Bill Russell (San Francisco)" Then
        MsgBox "Bill Russel had a Final Four high 27 rebounds against Iowa on March 23, 1956"
    End If
    If N = "Mark Wade, (UNLV)" Then
        MsgBox "Mark Wade had 18 assists against Indiana, a Final Four record"
    End If
    If N = "Danny Manning (Kansas)" Then
        MsgBox "Danny Manning had 6 blocked shots versus Duke, a Final Four record"
    End If
    If N = "Tommy Amaker (Duke)" Then
        MsgBox " Tommy Amaker had 7 steals against Duke, tying a Final Four record with Mookie Blaylock"
    End If
    If N = "Mookie Blaylock (Oklahoma)" Then
        MsgBox "Mookie Blaylock had 7 steals against Oklahoma, tying a Final Four record with Tommy Amaker"
    End If
    If N = "Oscar Robertson (Cincinnati)" Then
        MsgBox "Oscar Robertson has one of two Triple-Doubles ever to be recorded in the Final Four with 39 points, 17 rebounds and 10 assists"
    End If
    If N = "Magic Johnson (Michigan St.)" Then
        MsgBox " Magic Johnson holds one of only two Triple-Doubles ever recorded in the Final Four with 29 points, 10 rebounds, and 10 assists"
    End If
    'reads input and gives information on player
    
End Sub

Private Sub picteam_Click()
    N = InputBox("Please Input a team name, you must spell it correctly", "Team")
    If N = "UCLA" Then
        MsgBox "UCLA has appeared in the tournament 36 times(3rd), won 80 games(4th), and has 11 NCAA Championships(1st)", , "Output"
    End If
    If N = "Indiana" Then
        MsgBox "Indiana has appeared 32 times(5th) won 58 games(6th), and has 5 NCAA Championships(3rd)", , "Output"
    End If
    If N = "North Carolina" Then
        MsgBox "North Carolina has appeared in the tournament 37 times(2nd), won 88 games(2nd), and has 4 NCAA Championships(4th)", , "Output"
    End If
    If N = "Louisville" Then
        MsgBox "Louisville has appeared in the tournament 31 times(6th)", , "Output"
    End If
    If N = "Kansas" Then
        MsgBox "Kansas has appeared in the tournament 34 times(4th), and won 73 games(5th)", , "Output"
    End If
    If N = "Kentucky" Then
        MsgBox "Kentucky has appeared in the tournament 46 times(2nd), won 96 games(1st), and has 7 NCAA Championships(2nd)", , "Output"
    End If
    If N = "Duke" Then
        MsgBox "Duke has won 83 games(3rd), and has 3 NCAA Championships(5th)", , "Output"
    End If
End Sub
