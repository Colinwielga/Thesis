VERSION 5.00
Begin VB.Form frmFavreGame
   BackColor       =   &H00800080&
   Caption         =   "The Favre Game"
   ClientHeight    =   10845
   ClientLeft      =   3540
   ClientTop       =   1320
   ClientWidth     =   15780
   LinkTopic       =   "Form1"
   Picture         =   "FavreGame.frx":0000
   ScaleHeight     =   10845
   ScaleWidth      =   15780
   Begin VB.CommandButton cmdPlay3
      BackColor       =   &H000000FF&
      Caption         =   "Play 3"
      BeginProperty Font
         Name            =   "Franklin Gothic Demi"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   7
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton cmdPlay4
      BackColor       =   &H000000FF&
      Caption         =   "Play 4"
      BeginProperty Font
         Name            =   "Franklin Gothic Demi"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   5
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton cmdPlay5
      BackColor       =   &H000000FF&
      Caption         =   "Play 5"
      BeginProperty Font
         Name            =   "Franklin Gothic Demi"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton cmdPlay2
      BackColor       =   &H000000FF&
      Caption         =   "Play 2"
      BeginProperty Font
         Name            =   "Franklin Gothic Demi"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   2055
   End
   Begin VB.PictureBox picResults
      Height          =   2175
      Left            =   13920
      ScaleHeight     =   2115
      ScaleWidth      =   1635
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmdPlay1
      BackColor       =   &H000000FF&
      Caption         =   "Play 1"
      BeginProperty Font
         Name            =   "Franklin Gothic Demi"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton cmdStart
      BackColor       =   &H00800080&
      Caption         =   "Play the Brett Favre Game!"
      BeginProperty Font
         Name            =   "Franklin Gothic Demi"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   11175
   End
   Begin VB.CommandButton cmdMenu
      BackColor       =   &H0000FFFF&
      Caption         =   "Return to Main Menu ==>"
      BeginProperty Font
         Name            =   "Franklin Gothic Demi"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9360
      Width           =   2775
   End
End
Attribute VB_Name = "frmFavreGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Brett Favre Experience
'frmFavreGame
'Doug Donaldson
'2/24/10

'this form is strictly about fun. it simulates in-game decisions brett would have to make, and gives the user a score based on the decisions they made.



'dim local variables
Dim PlayCTR As Integer, Time As Integer, Run As Integer, Pass As Integer, Play As Integer, YardsLeft As Integer, QBGameRating As Single
Dim GameAttempts As Integer, GameCompletions As Integer, PassingYards As Integer, PassTDs As Integer, Interceptions As Integer


Private Sub cmdStart_Click()

'remove start button so that user can begin game
cmdStart.Visible = False
cmdPlay1.Visible = True

'initialize variables
Run = 16
Pass = 7
Time = 112
PlayCTR = 10
YardsLeft = 80
GameAttempts = 0
GameCompletions = 0
PassingYards = 0
PassTDs = 0
Interceptions = 0

'present game instruction/goal to user
MsgBox "You have 5 plays or 1 minute and 52 seconds to drive 80 yards down the field and score a touchdown to win.", , ""

End Sub


Private Sub cmdPlay1_Click()

picResults.Cls

        Play = InputBox("Choose a play: Run left = 1, Run middle = 2, Run right = 3, Pass deep = 4, Pass over the middle = 5, Pass left = 6, Pass right = 7")

            Select Case Play
            Case 1

                PlayCTR = PlayCTR - 1
                Time = Time - Run
                YardsLeft = YardsLeft - 4
                MsgBox "You gained 4 yards!"
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


            Case 2

                PlayCTR = PlayCTR - 1
                Time = Time - Run
                YardsLeft = YardsLeft - 3
                MsgBox "You only gained 3 yards."
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


           Case 3

                PlayCTR = PlayCTR - 1
                Time = Time - Run
                YardsLeft = YardsLeft - 6
                MsgBox "You gained 6 yards!"
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


            Case 4

                PlayCTR = PlayCTR - 1
                Time = Time - Pass
                YardsLeft = YardsLeft + 5
                GameAttempts = GameAttempts + 1
                GameCompletions = GameCompletions + 0
                PassingYards = PassingYards + 0
                PassTDs = PassTDs + 0
                Interceptions = Interceptions + 0
                MsgBox "In order to protect you, your lineman had to commit a penalty. Results in a 5 yard loss."
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


            Case 5

                PlayCTR = PlayCTR - 1
                Time = Time - Pass
                YardsLeft = YardsLeft - 7
                GameAttempts = GameAttempts + 1
                GameCompletions = GameCompletions + 1
                PassingYards = PassingYards + 7
                PassTDs = PassTDs + 0
                Interceptions = Interceptions + 0
                MsgBox "Thrown over the middle for a 7 yard gain!"
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


            Case 6

                PlayCTR = PlayCTR - 1
                Time = Time - Pass
                YardsLeft = YardsLeft - 4
                GameAttempts = GameAttempts + 1
                GameCompletions = GameCompletions + 1
                PassingYards = PassingYards + 4
                PassTDs = PassTDs + 0
                Interceptions = Interceptions + 0
                MsgBox "Thrown to the left for a minimal 4 yard gain."
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


            Case 7

                PlayCTR = PlayCTR - 1
                Time = Time - Pass
                YardsLeft = YardsLeft
                GameAttempts = GameAttempts + 1
                GameCompletions = GameCompletions + 0
                PassingYards = PassingYards + 0
                PassTDs = PassTDs + 0
                Interceptions = Interceptions + 0
                MsgBox "Ball was knocked down. No gain."
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


        End Select

End Sub

Private Sub cmdPlay2_Click(Index As Integer)

picResults.Cls

        Play = InputBox("Choose a play: Run left = 1, Run middle = 2, Run right = 3, Pass deep = 4, Pass over the middle = 5, Pass left = 6, Pass right = 7")

            Select Case Play
            Case 1

                PlayCTR = PlayCTR - 1
                Time = Time - Run
                YardsLeft = YardsLeft + 1
                MsgBox "You lost 1 yard!"
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


            Case 2

                PlayCTR = PlayCTR - 1
                Time = Time - Run
                YardsLeft = YardsLeft - 3
                MsgBox "You only gained 3 yards."
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


           Case 3

                PlayCTR = PlayCTR - 1
                Time = Time - Run
                YardsLeft = YardsLeft - 20
                MsgBox "You gained 20 yards!"
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


            Case 4

                PlayCTR = PlayCTR - 1
                Time = Time - Pass
                YardsLeft = YardsLeft - 15
                GameAttempts = GameAttempts + 1
                GameCompletions = GameCompletions + 1
                PassingYards = PassingYards + 15
                PassTDs = PassTDs + 0
                Interceptions = Interceptions + 0
                MsgBox "You gained 15 yards on that play."
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


            Case 5

                PlayCTR = PlayCTR - 1
                Time = Time - Pass
                YardsLeft = YardsLeft - 0
                GameAttempts = GameAttempts + 1
                GameCompletions = GameCompletions + 0
                PassingYards = PassingYards + 0
                PassTDs = PassTDs + 0
                Interceptions = Interceptions + 1
                QBGameRating = (((((GameCompletions / GameAttempts) * 100) - 30) / 20) + (((PassTDs / GameAttempts) * 100) / 5) + ((9.5 - ((Interceptions / GameAttempts) * 100)) / 4) + (((PassingYards / GameAttempts) - 3) / 4) / 0.06)
                MsgBox "Thrown over the middle for an interception! You have lost the game with a QB Rating of " & QBGameRating
                cmdStart.Visible = True

            Case 6

                PlayCTR = PlayCTR - 1
                Time = Time - Pass
                YardsLeft = YardsLeft - 8
                GameAttempts = GameAttempts + 1
                GameCompletions = GameCompletions + 1
                PassingYards = PassingYards + 8
                PassTDs = PassTDs + 0
                Interceptions = Interceptions + 0
                MsgBox "Thrown to the left for an 8 yard gain."
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


            Case 7

                PlayCTR = PlayCTR - 1
                Time = Time - Pass
                YardsLeft = YardsLeft - 0
                GameAttempts = GameAttempts + 1
                GameCompletions = GameCompletions + 1
                PassingYards = PassingYards + YardsLeft
                PassTDs = PassTDs + 1
                Interceptions = Interceptions + 0
                QBGameRating = (((((GameCompletions / GameAttempts) * 100) - 30) / 20) + (((PassTDs / GameAttempts) * 100) / 5) + ((9.5 - ((Interceptions / GameAttempts) * 100)) / 4) + (((PassingYards / GameAttempts) - 3) / 4) / 0.06)
                MsgBox "Touchdown! You have won the game with a Quarterback Rating of: " & QBGameRating
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR
                cmdStart.Visible = True

        End Select

End Sub

Private Sub cmdPlay3_Click(Index As Integer)

picResults.Cls

        Play = InputBox("Choose a play: Run left = 1, Run middle = 2, Run right = 3, Pass deep = 4, Pass over the middle = 5, Pass left = 6, Pass right = 7")

            Select Case Play
            Case 1

                PlayCTR = PlayCTR - 1
                Time = Time - Run
                YardsLeft = 0
                MsgBox "You scored a touchdown!"
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


            Case 2

                PlayCTR = PlayCTR - 1
                Time = Time - Run
                YardsLeft = YardsLeft - 1
                MsgBox "You only gained 1 yard."
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


           Case 3

                PlayCTR = PlayCTR - 1
                Time = Time - Run
                YardsLeft = YardsLeft - 4
                MsgBox "You gained 4 yards!"
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


            Case 4

                PlayCTR = PlayCTR - 1
                Time = Time - Pass
                YardsLeft = YardsLeft - 9
                GameAttempts = GameAttempts + 1
                GameCompletions = GameCompletions + 1
                PassingYards = PassingYards + 9
                PassTDs = PassTDs + 0
                Interceptions = Interceptions + 0
                MsgBox "Your receiver caught the ball for a nine yard gain!"
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


            Case 5

                PlayCTR = PlayCTR - 1
                Time = Time - Pass
                YardsLeft = YardsLeft - 7
                GameAttempts = GameAttempts + 1
                GameCompletions = GameCompletions + 1
                PassingYards = PassingYards + 7
                PassTDs = PassTDs + 0
                Interceptions = Interceptions + 0
                MsgBox "Thrown over the middle for a 7 yard gain!"
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


            Case 6

                PlayCTR = PlayCTR - 1
                Time = Time - Pass
                YardsLeft = YardsLeft - 0
                GameAttempts = GameAttempts + 1
                GameCompletions = GameCompletions + 1
                PassingYards = PassingYards + 0
                PassTDs = PassTDs + 0
                Interceptions = Interceptions + 0
                MsgBox "Thrown to the left for no gain."
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


            Case 7

                PlayCTR = PlayCTR - 1
                Time = Time - Pass
                YardsLeft = YardsLeft - 6
                GameAttempts = GameAttempts + 1
                GameCompletions = GameCompletions + 1
                PassingYards = PassingYards + 6
                PassTDs = PassTDs + 0
                Interceptions = Interceptions + 0
                MsgBox "Ball was caught for a six yard gain."
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


        End Select

End Sub

Private Sub cmdPlay4_Click(Index As Integer)

picResults.Cls

        Play = InputBox("Choose a play: Run left = 1, Run middle = 2, Run right = 3, Pass deep = 4, Pass over the middle = 5, Pass left = 6, Pass right = 7")

            Select Case Play
            Case 1

                PlayCTR = PlayCTR - 1
                Time = Time - Run
                YardsLeft = YardsLeft - 4
                MsgBox "You gained 4 yards!"
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


            Case 2

                PlayCTR = PlayCTR - 1
                Time = Time - Run
                YardsLeft = YardsLeft - 23
                MsgBox "You gained 23 yards!"
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


           Case 3

                PlayCTR = PlayCTR - 1
                Time = Time - Run
                YardsLeft = YardsLeft + 6
                MsgBox "You lost 6 yards!"
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


            Case 4

                PlayCTR = PlayCTR - 1
                Time = Time - Pass
                YardsLeft = YardsLeft - 7
                GameAttempts = GameAttempts + 1
                GameCompletions = GameCompletions + 1
                PassingYards = PassingYards + 7
                PassTDs = PassTDs + 0
                Interceptions = Interceptions + 0
                MsgBox "A completed pass to the running back goes for a seven yard gain."
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


            Case 5

                PlayCTR = PlayCTR - 1
                Time = Time - Pass
                YardsLeft = YardsLeft - 7
                GameAttempts = GameAttempts + 1
                GameCompletions = GameCompletions + 1
                PassingYards = PassingYards + 7
                PassTDs = PassTDs + 0
                Interceptions = Interceptions + 0
                MsgBox "Thrown over the middle for a 7 yard gain!"
                picResults.Print "Time left: "; Time; " seconds."
                picResults.Print "Yards left: "; YardsLeft
                picResults.Print "Plays left: "; PlayCTR


            Case 6

                PlayCTR = PlayCTR - 1
                Time = Time - Pass
                YardsLeft = YardsLeft + 0
                GameAttempts = GameAttempts + 1
                GameCompletions = GameCompletions + 1
                PassingYards = PassingYards + YardsLeft
                PassTDs = PassTDs + 1
                Interceptions = Interceptions + 0
                QBGameRating = (((((GameCompletions / GameAttempts) * 100) - 30) / 20) + (((PassTDs / GameAttempts) * 100) / 5) + ((9.5 - ((Interceptions / GameAttempts) * 100)) / 4) + (((PassingYards / GameAttempts) - 3) / 4) / 0.06)
                MsgBox "Thrown to the left for a touchdown! You have won the game with a QB Rating of " & QBGameRating


            Case 7

                PlayCTR = PlayCTR - 1
                Time = Time - Pass
                YardsLeft = YardsLeft
                GameAttempts = GameAttempts + 1
                GameCompletions = GameCompletions + 0
                PassingYards = PassingYards + 0
                PassTDs = PassTDs + 0
                Interceptions = Interceptions + 1
                QBGameRating = (((((GameCompletions / GameAttempts) * 100) - 30) / 20) + (((PassTDs / GameAttempts) * 100) / 5) + ((9.5 - ((Interceptions / GameAttempts) * 100)) / 4) + (((PassingYards / GameAttempts) - 3) / 4) / 0.06)
                MsgBox "Ball was tipped and intercepted. Game over. QB Rating of " & QBGameRating


        End Select

End Sub

Private Sub cmdPlay5_Click(Index As Integer)

picResults.Cls

        Play = InputBox("Choose a play: Run left = 1, Run middle = 2, Run right = 3, Pass deep = 4, Pass over the middle = 5, Pass left = 6, Pass right = 7")

            Select Case Play
            Case 1

                PlayCTR = PlayCTR - 1
                Time = Time - Run
                YardsLeft = YardsLeft
                MsgBox "You fumbled the ball. Game over!"
                cmdStart.Visible = True

            Case 2

                PlayCTR = PlayCTR - 1
                Time = Time - Run
                YardsLeft = 0
                MsgBox "You ran for a touchdown! Game over."
                cmdStart.Visible = True

           Case 3

                PlayCTR = PlayCTR - 1
                Time = Time - Run
                YardsLeft = YardsLeft
                MsgBox "You broke your leg! Game over. "
                cmdStart.Visible = True


            Case 4
                PlayCTR = PlayCTR - 1
                Time = Time - Pass
                YardsLeft = YardsLeft
                GameAttempts = GameAttempts
                GameCompletions = GameCompletions
                PassingYards = PassingYards
                PassTDs = PassTDs
                Interceptions = Interceptions
                MsgBox "The other team has forfeited due to their intimidation from your abilities. Thanks for playing oh Great One."
                cmdStart.Visible = True

            Case 5
                PlayCTR = PlayCTR - 1
                Time = Time - Pass
                YardsLeft = YardsLeft
                GameAttempts = GameAttempts
                GameCompletions = GameCompletions
                PassingYards = PassingYards
                PassTDs = PassTDs + 0
                Interceptions = Interceptions + 0
                MsgBox "Aliens have attacked and taken you hostage. Game over."
                cmdStart.Visible = True

            Case 6
                PlayCTR = PlayCTR - 1
                Time = Time - Pass
                YardsLeft = YardsLeft
                GameAttempts = GameAttempts + 1
                GameCompletions = GameCompletions + 1
                PassingYards = PassingYards + YardsLeft
                PassTDs = PassTDs + 1
                Interceptions = Interceptions + 0
                QBGameRating = (((((GameCompletions / GameAttempts) * 100) - 30) / 20) + (((PassTDs / GameAttempts) * 100) / 5) + ((9.5 - ((Interceptions / GameAttempts) * 100)) / 4) + (((PassingYards / GameAttempts) - 3) / 4) / 0.06)
                MsgBox "You have thrown a touchdown pass! You win!"
                cmdStart.Visible = True

            Case 7
                PlayCTR = PlayCTR - 1
                Time = Time - Pass
                YardsLeft = YardsLeft
                GameAttempts = GameAttempts
                GameCompletions = GameCompletions + 0
                PassingYards = PassingYards + 0
                PassTDs = PassTDs + 0
                Interceptions = Interceptions + 0
                MsgBox "Earthquake strikes the stadium. Lights out on this one."
                cmdStart.Visible = True

        End Select

End Sub


Private Sub cmdMenu_Click()
frmFavreGame.Hide
frmMain.Show
End Sub


