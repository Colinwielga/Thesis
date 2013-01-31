VERSION 5.00
Begin VB.Form frmRoommateChallenge3
   BackColor       =   &H00000000&
   Caption         =   "Roommate Challenge"
   ClientHeight    =   13005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   ScaleHeight     =   13005
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpinion
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to hear what the judges have to say about your score!"
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8880
      Width           =   2775
   End
   Begin VB.CommandButton Command2
      BackColor       =   &H00FFC0C0&
      Caption         =   "Quit"
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   11160
      Width           =   1335
   End
   Begin VB.CommandButton Command1
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to see other teams scores and determine how you compare!"
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8880
      Width           =   2775
   End
   Begin VB.CommandButton cmdFinalRound
      BackColor       =   &H00C000C0&
      Caption         =   "Final Round"
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8040
      Width           =   2775
   End
   Begin VB.CommandButton cmdRound2
      BackColor       =   &H00FF00FF&
      Caption         =   "Round 2"
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7320
      Width           =   2775
   End
   Begin VB.CommandButton cmdRound1
      BackColor       =   &H00FF80FF&
      Caption         =   "Round 1 "
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6480
      Width           =   2775
   End
   Begin VB.CommandButton cmdPlay
      BackColor       =   &H00FFC0FF&
      Caption         =   "Click here to Start!"
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5400
      Width           =   1455
   End
   Begin VB.PictureBox picPlayer4Name
      BackColor       =   &H00FFFFFF&
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      ScaleHeight     =   555
      ScaleWidth      =   1635
      TabIndex        =   10
      Top             =   1920
      Width           =   1695
   End
   Begin VB.PictureBox picPlayer3Name
      BackColor       =   &H00FFFFFF&
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      ScaleHeight     =   555
      ScaleWidth      =   1635
      TabIndex        =   9
      Top             =   1920
      Width           =   1695
   End
   Begin VB.PictureBox picPlayer2Name
      BackColor       =   &H00FFFFFF&
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   1635
      TabIndex        =   8
      Top             =   1920
      Width           =   1695
   End
   Begin VB.PictureBox picPlayer1Name
      BackColor       =   &H00FFFFFF&
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      ScaleHeight     =   435
      ScaleWidth      =   1635
      TabIndex        =   7
      Top             =   1920
      Width           =   1695
   End
   Begin VB.PictureBox picTeam2Results
      BackColor       =   &H0000FFFF&
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      ScaleHeight     =   795
      ScaleWidth      =   3555
      TabIndex        =   5
      Top             =   4320
      Width           =   3615
   End
   Begin VB.PictureBox picTeam1Results
      BackColor       =   &H0000FFFF&
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      ScaleHeight     =   915
      ScaleWidth      =   3555
      TabIndex        =   4
      Top             =   4200
      Width           =   3615
   End
   Begin VB.PictureBox picPlayer3
      Height          =   1455
      Left            =   6960
      ScaleHeight     =   1395
      ScaleWidth      =   1275
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.PictureBox picPlayer1
      Height          =   1455
      Left            =   480
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.PictureBox picPlayer2
      Height          =   1455
      Left            =   2400
      ScaleHeight     =   1395
      ScaleWidth      =   1275
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin VB.PictureBox picPlayer4
      Height          =   1455
      Left            =   8760
      ScaleHeight     =   1395
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label LBLversus
      BackColor       =   &H00000000&
      Caption         =   "VS"
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   72
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1695
      Left            =   4320
      TabIndex        =   19
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblteam2
      BackColor       =   &H00000000&
      Caption         =   "Team 2"
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7920
      TabIndex        =   12
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblTeam1
      BackColor       =   &H00000000&
      Caption         =   "Team 1"
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1560
      TabIndex        =   11
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label LblTitle
      BackColor       =   &H00FF80FF&
      Caption         =   "     The Roommate Challenge!"
      BeginProperty Font
         Name            =   "Wide Latin"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   240
      Width           =   11895
   End
End
Attribute VB_Name = "frmRoommateChallenge3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



    ' a comment for EVERY new line!
                Dim team1Total As Integer
    ' a comment for EVERY new line!
                Dim team2Total As Integer
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                Private Sub cmdFinalRound_Click()
    ' a comment for EVERY new line!
                Dim AnswerF11 As String, AnswerF12 As String, AnswerF13 As String, AnswerF14 As String
    ' a comment for EVERY new line!
                Dim AnswerF21 As String, AnswerF22 As String, AnswerF23 As String, AnswerF24 As String
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                MsgBox "This round is worth 10 points per question!"
    ' a comment for EVERY new line!
                MsgBox RCPlayer2Name & " and " & RCPlayer4Name & " please leave or look away."
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                AnswerF11 = InputBox("What is your roommate most afraid of?", RCPlayer1Name & " answer this question.")
    ' a comment for EVERY new line!
                AnswerF13 = InputBox("What is your roommate most afraid of?", RCPlayer3Name & " answer this question.")
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                MsgBox "Now " & RCPlayer2Name & " and " & RCPlayer4Name & " come back to answer the next question!"
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                AnswerF12 = InputBox("When asked the question: What is your roommate most afraid of? Your roommate answered what?", RCPlayer2Name & " answer this question")
    ' a comment for EVERY new line!
                If LCase(AnswerF11) = LCase(AnswerF12) Then
    ' a comment for EVERY new line!
                MsgBox "Correct!"
    ' a comment for EVERY new line!
                team1Total = team1Total + 10
    ' a comment for EVERY new line!
                picTeam1Results.Cls
    ' a comment for EVERY new line!
                picTeam1Results.Print team1Total
    ' a comment for EVERY new line!
                Else
    ' a comment for EVERY new line!
                MsgBox "No your roommate answered " & AnswerF11
    ' a comment for EVERY new line!
                End If
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                AnswerF14 = InputBox("When asked the question: What is your roommate most afraid of? Your roommate answered what?", RCPlayer4Name & " answer this question")
    ' a comment for EVERY new line!
                If LCase(AnswerF14) = LCase(AnswerF13) Then
    ' a comment for EVERY new line!
                MsgBox "Correct!"
    ' a comment for EVERY new line!
                team2Total = team2Total + 10
    ' a comment for EVERY new line!
                picTeam2Results.Cls
    ' a comment for EVERY new line!
                picTeam2Results.Print team2Total
    ' a comment for EVERY new line!
                Else
    ' a comment for EVERY new line!
                MsgBox "No your roommate answered " & AnswerF13
    ' a comment for EVERY new line!
                End If
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                'second question. Player 2 and 4 answer first set.  1 and 3 answer second set.
    ' a comment for EVERY new line!
                MsgBox RCPlayer1Name & " and " & RCPlayer3Name & " please leave or look away."
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                AnswerF22 = InputBox("What is your roommate's ideal first date: A.Romantic Dinner B.Movie C.Concert D.Long Walk on the beach?", RCPlayer2Name & " answer this question.")
    ' a comment for EVERY new line!
                AnswerF24 = InputBox("What is your roommate's ideal first date: A.Romantic Dinner B.Movie C.Concert D.Long Walk on the beach?", RCPlayer4Name & " answer this question.")
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                MsgBox "Now " & RCPlayer1Name & " and " & RCPlayer3Name & " come back to answer the next question!"
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                AnswerF21 = InputBox("When asked the question: What is your roommate's ideal first date: A.Romantic Dinner B.Movie C.Concert D.Long Walk on the beach? Your roommate answered what?", RCPlayer1Name & " answer this question")
    ' a comment for EVERY new line!
                If LCase(AnswerF21) = LCase(AnswerF22) Then
    ' a comment for EVERY new line!
                MsgBox "Correct!"
    ' a comment for EVERY new line!
                team1Total = team1Total + 10
    ' a comment for EVERY new line!
                picTeam1Results.Cls
    ' a comment for EVERY new line!
                picTeam1Results.Print team1Total
    ' a comment for EVERY new line!
                Else
    ' a comment for EVERY new line!
                MsgBox "No your roommate answered " & AnswerF22
    ' a comment for EVERY new line!
                End If
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                AnswerF23 = InputBox("When asked the question: What is your roommate's ideal first date: A.Romantic Dinner B.Movie C.Concert D.Long Walk on the beach? Your roommate answered what?", RCPlayer3Name & " answer this question")
    ' a comment for EVERY new line!
                If LCase(AnswerF24) = LCase(AnswerF23) Then
    ' a comment for EVERY new line!
                MsgBox "Correct!"
    ' a comment for EVERY new line!
                team2Total = team2Total + 10
    ' a comment for EVERY new line!
                picTeam2Results.Cls
    ' a comment for EVERY new line!
                picTeam2Results.Print team2Total
    ' a comment for EVERY new line!
                Else
    ' a comment for EVERY new line!
                MsgBox "No your roommate answered " & AnswerF24
    ' a comment for EVERY new line!
                End If
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                'declare a winner
    ' a comment for EVERY new line!
                If team2Total = team1Total Then
    ' a comment for EVERY new line!
                MsgBox "It's a tie!"
    ' a comment for EVERY new line!
                ElseIf team2Total > team1Total Then
    ' a comment for EVERY new line!
                MsgBox "Team 2 wins! Congrats " & RCPlayer3Name & " & " & RCPlayer4Name & "!", , "WINNERS"
    ' a comment for EVERY new line!
                Else
    ' a comment for EVERY new line!
                MsgBox "Team 1 wins! Congrats " & RCPlayer1Name & " & " & RCPlayer2Name & "!", , "WINNERS"
    ' a comment for EVERY new line!
                End If
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                End Sub
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!

    ' a comment for EVERY new line!

    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                Private Sub cmdOpinion_Click()
    ' a comment for EVERY new line!
                frmRoommateChallenge3.Hide
    ' a comment for EVERY new line!
                frmJudges.Show
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                End Sub
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                Private Sub cmdPlay_Click()
    ' a comment for EVERY new line!
                'load the pictures and names of the players
    ' a comment for EVERY new line!
                'player 1
    ' a comment for EVERY new line!
                picPlayer1Name.Print RCPlayer1Name
    ' a comment for EVERY new line!
                If RCPlayer1Pic = 1 Then
    ' a comment for EVERY new line!
                picPlayer1.Picture = LoadPicture(App.Path & "\Caitlin.bmp")
    ' a comment for EVERY new line!
                ElseIf RCPlayer1Pic = 2 Then
    ' a comment for EVERY new line!
                picPlayer1.Picture = LoadPicture(App.Path & "\Emily.bmp")
    ' a comment for EVERY new line!
                ElseIf RCPlayer1Pic = 3 Then
    ' a comment for EVERY new line!
                picPlayer1.Picture = LoadPicture(App.Path & "\me.bmp")
    ' a comment for EVERY new line!
                ElseIf RCPlayer1Pic = 4 Then
    ' a comment for EVERY new line!
                picPlayer1.Picture = LoadPicture(App.Path & "\heather.bmp")
    ' a comment for EVERY new line!
                ElseIf RCPlayer1Pic = 5 Then
    ' a comment for EVERY new line!
                picPlayer1.Picture = LoadPicture(App.Path & "\katie.bmp")
    ' a comment for EVERY new line!
                Else: RCPlayer1Pic = 6
    ' a comment for EVERY new line!
                picPlayer1.Picture = LoadPicture(App.Path & "\steph.bmp")
    ' a comment for EVERY new line!
                End If
    ' a comment for EVERY new line!
                'player 2
    ' a comment for EVERY new line!
                picPlayer2Name.Print RCPlayer2Name
    ' a comment for EVERY new line!
                If RCPlayer2Pic = 1 Then
    ' a comment for EVERY new line!
                picPlayer2.Picture = LoadPicture(App.Path & "\Caitlin.bmp")
    ' a comment for EVERY new line!
                ElseIf RCPlayer2Pic = 2 Then
    ' a comment for EVERY new line!
                picPlayer2.Picture = LoadPicture(App.Path & "\Emily.bmp")
    ' a comment for EVERY new line!
                ElseIf RCPlayer2Pic = 3 Then
    ' a comment for EVERY new line!
                picPlayer2.Picture = LoadPicture(App.Path & "\me.bmp")
    ' a comment for EVERY new line!
                ElseIf RCPlayer2Pic = 4 Then
    ' a comment for EVERY new line!
                picPlayer2.Picture = LoadPicture(App.Path & "\heather.bmp")
    ' a comment for EVERY new line!
                ElseIf RCPlayer2Pic = 5 Then
    ' a comment for EVERY new line!
                picPlayer2.Picture = LoadPicture(App.Path & "\katie.bmp")
    ' a comment for EVERY new line!
                Else: RCPlayer2Pic = 6
    ' a comment for EVERY new line!
                picPlayer2.Picture = LoadPicture(App.Path & "\steph.bmp")
    ' a comment for EVERY new line!
                End If
    ' a comment for EVERY new line!
                'player 3
    ' a comment for EVERY new line!
                picPlayer3Name.Print RCPlayer3Name
    ' a comment for EVERY new line!
                If RCPlayer3Pic = 1 Then
    ' a comment for EVERY new line!
                picPlayer3.Picture = LoadPicture(App.Path & "\Caitlin.bmp")
    ' a comment for EVERY new line!
                ElseIf RCPlayer3Pic = 2 Then
    ' a comment for EVERY new line!
                picPlayer3.Picture = LoadPicture(App.Path & "\Emily.bmp")
    ' a comment for EVERY new line!
                ElseIf RCPlayer3Pic = 3 Then
    ' a comment for EVERY new line!
                picPlayer3.Picture = LoadPicture(App.Path & "\me.bmp")
    ' a comment for EVERY new line!
                ElseIf RCPlayer3Pic = 4 Then
    ' a comment for EVERY new line!
                picPlayer3.Picture = LoadPicture(App.Path & "\heather.bmp")
    ' a comment for EVERY new line!
                ElseIf RCPlayer3Pic = 5 Then
    ' a comment for EVERY new line!
                picPlayer3.Picture = LoadPicture(App.Path & "\katie.bmp")
    ' a comment for EVERY new line!
                Else: RCPlayer3Pic = 6
    ' a comment for EVERY new line!
                picPlayer3.Picture = LoadPicture(App.Path & "\steph.bmp")
    ' a comment for EVERY new line!
                End If
    ' a comment for EVERY new line!
                'player 4
    ' a comment for EVERY new line!
                picPlayer4Name.Print RCPlayer4Name
    ' a comment for EVERY new line!
                If RCPlayer4Pic = 1 Then
    ' a comment for EVERY new line!
                picPlayer4.Picture = LoadPicture(App.Path & "\Caitlin.bmp")
    ' a comment for EVERY new line!
                ElseIf RCPlayer4Pic = 2 Then
    ' a comment for EVERY new line!
                picPlayer4.Picture = LoadPicture(App.Path & "\Emily.bmp")
    ' a comment for EVERY new line!
                ElseIf RCPlayer4Pic = 3 Then
    ' a comment for EVERY new line!
                picPlayer4.Picture = LoadPicture(App.Path & "\me.bmp")
    ' a comment for EVERY new line!
                ElseIf RCPlayer4Pic = 4 Then
    ' a comment for EVERY new line!
                picPlayer4.Picture = LoadPicture(App.Path & "\heather.bmp")
    ' a comment for EVERY new line!
                ElseIf RCPlayer4Pic = 5 Then
    ' a comment for EVERY new line!
                picPlayer4.Picture = LoadPicture(App.Path & "\katie.bmp")
    ' a comment for EVERY new line!
                Else: RCPlayer4Pic = 6
    ' a comment for EVERY new line!
                picPlayer4.Picture = LoadPicture(App.Path & "\steph.bmp")
    ' a comment for EVERY new line!
                End If
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                MsgBox RCPlayer2Name & " and " & RCPlayer4Name & " need to leave the room or at least turn away from the screen."
    ' a comment for EVERY new line!
                MsgBox "Now " & RCPlayer1Name & " and " & RCPlayer3Name & " start by clicking the Round 1 button!"
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                End Sub
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                Private Sub cmdRound1_Click()
    ' a comment for EVERY new line!
                Dim Answer11 As String, Answer12 As String
    ' a comment for EVERY new line!
                Dim Answer13 As String, Answer14 As String
    ' a comment for EVERY new line!
                Dim Answer21 As String, Answer22 As String
    ' a comment for EVERY new line!
                Dim Answer23 As String, Answer24 As String
    ' a comment for EVERY new line!
                Dim Answer31 As String, Answer32 As String
    ' a comment for EVERY new line!
                Dim Answer33 As String, Answer34 As String
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                team1Total = 0
    ' a comment for EVERY new line!
                team2Total = 0
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                Answer11 = InputBox("If your roomate had to eat one thing for the rest of her life what would it be?", RCPlayer1Name & " answer this question")
    ' a comment for EVERY new line!
                Answer13 = InputBox("If your roomate had to eat one thing for the rest of her life what would it be?", RCPlayer3Name & " answer this question")
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                Answer21 = InputBox("What color is your roommate's favorite shirt?", RCPlayer1Name & " answer this question")
    ' a comment for EVERY new line!
                Answer23 = InputBox("What color is your roommate's favorite shirt?", RCPlayer3Name & " answer this question")
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                Answer31 = InputBox("If your roomate was stranded on an island she would A. Build an incredible raft and escape; B. Panic and scream; C. Spend the time lying on the beach soaking up the sun saying 'Hey at least I'll get a killer tan'?", RCPlayer1Name & " answer this question")
    ' a comment for EVERY new line!
                Answer33 = InputBox("If your roomate was stranded on an island she would A. Build an incredible raft and escape; B. Panic and scream; C. Spend the time lying on the beach soaking up the sun saying 'Hey at least I'll get a killer tan'?", RCPlayer3Name & " answer this question")
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                MsgBox "Now " & RCPlayer2Name & " and " & RCPlayer4Name & " come back to answer the next questions!"
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                'team1 ?1
    ' a comment for EVERY new line!
                Answer12 = InputBox("When asked the question: If your roomate had to eat one thing for the rest of her life what would it be? Your roommate answered what?", RCPlayer2Name & " answer this question")
    ' a comment for EVERY new line!
                If LCase(Answer11) = LCase(Answer12) Then
    ' a comment for EVERY new line!
                MsgBox "Correct!"
    ' a comment for EVERY new line!
                team1Total = team1Total + 5
    ' a comment for EVERY new line!
                picTeam1Results.Print team1Total
    ' a comment for EVERY new line!
                Else
    ' a comment for EVERY new line!
                MsgBox "No your roommate answered " & Answer11
    ' a comment for EVERY new line!
                picTeam1Results.Print team1Total
    ' a comment for EVERY new line!
                End If
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                'team2 ?2
    ' a comment for EVERY new line!
                Answer14 = InputBox("When asked the question: If your roomate had to eat one thing for the rest of her life what would it be? Your roommate answered what?", RCPlayer4Name & " answer this question") 'team 2 scoring
    ' a comment for EVERY new line!
                If LCase(Answer13) = LCase(Answer14) Then
    ' a comment for EVERY new line!
                MsgBox "Correct!"
    ' a comment for EVERY new line!
                team2Total = team2Total + 5
    ' a comment for EVERY new line!
                picTeam2Results.Print team2Total
    ' a comment for EVERY new line!
                Else
    ' a comment for EVERY new line!
                MsgBox "No your roommate answered " & Answer13
    ' a comment for EVERY new line!
                picTeam2Results.Print team2Total
    ' a comment for EVERY new line!
                End If
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                'team1 ?2
    ' a comment for EVERY new line!
                Answer22 = InputBox("When asked the question: What color is your roommate's favorite shirt? Your roommate answered what?", RCPlayer2Name & " answer this question")
    ' a comment for EVERY new line!
                If LCase(Answer21) = LCase(Answer22) Then
    ' a comment for EVERY new line!
                MsgBox "Correct!"
    ' a comment for EVERY new line!
                team1Total = team1Total + 5
    ' a comment for EVERY new line!
                picTeam1Results.Cls
    ' a comment for EVERY new line!
                picTeam1Results.Print team1Total
    ' a comment for EVERY new line!
                Else
    ' a comment for EVERY new line!
                MsgBox "No your roommate answered " & Answer21
    ' a comment for EVERY new line!
                End If
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                'team2 ?2
    ' a comment for EVERY new line!
                Answer24 = InputBox("When asked the question: What color is your roommate's favorite shirt? Your roommate answered what?", RCPlayer4Name & " answer this question")
    ' a comment for EVERY new line!
                If LCase(Answer23) = LCase(Answer24) Then
    ' a comment for EVERY new line!
                MsgBox "Correct!"
    ' a comment for EVERY new line!
                team2Total = team2Total + 5
    ' a comment for EVERY new line!
                picTeam2Results.Cls
    ' a comment for EVERY new line!
                picTeam2Results.Print team2Total
    ' a comment for EVERY new line!
                Else
    ' a comment for EVERY new line!
                MsgBox "No your roommate answered " & Answer23
    ' a comment for EVERY new line!
                End If
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                'team1 ?3
    ' a comment for EVERY new line!
                Answer32 = InputBox("When asked the question: If your roomate was stranded on an island she would A. Build an incredible raft and escape; B. Panic and scream; C. Spend the time lying on the beach soaking up the sun saying 'Hey at least I'll get a killer tan'? Your roommate answered what?", RCPlayer2Name & " answer this question")
    ' a comment for EVERY new line!
                If LCase(Answer31) = LCase(Answer32) Then
    ' a comment for EVERY new line!
                MsgBox "Correct!"
    ' a comment for EVERY new line!
                team1Total = team1Total + 5
    ' a comment for EVERY new line!
                picTeam1Results.Cls
    ' a comment for EVERY new line!
                picTeam1Results.Print team1Total
    ' a comment for EVERY new line!
                Else
    ' a comment for EVERY new line!
                MsgBox "No your roommate answered " & Answer31
    ' a comment for EVERY new line!
                End If
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                'team2 ?3
    ' a comment for EVERY new line!
                Answer34 = InputBox("When asked the question: If your roomate were stranded on an island she would A. Build an incredible raft and escape; B. Panic and scream; C. Spend the her time lying on the beach soaking up the sun saying 'Hey at least I'll get a killer tan'? Your roommate answered what?", RCPlayer4Name & " answer this question")
    ' a comment for EVERY new line!
                If LCase(Answer33) = LCase(Answer34) Then
    ' a comment for EVERY new line!
                MsgBox "Correct!"
    ' a comment for EVERY new line!
                team2Total = team2Total + 5
    ' a comment for EVERY new line!
                picTeam2Results.Cls
    ' a comment for EVERY new line!
                picTeam2Results.Print team2Total
    ' a comment for EVERY new line!
                Else
    ' a comment for EVERY new line!
                MsgBox "No your roommate answered " & Answer33
    ' a comment for EVERY new line!
                End If
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                End Sub
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                Private Sub cmdRound2_Click()
    ' a comment for EVERY new line!
                Dim Answer41 As String, Answer42 As String
    ' a comment for EVERY new line!
                Dim Answer43 As String, Answer44 As String
    ' a comment for EVERY new line!
                Dim Answer51 As String, Answer52 As String
    ' a comment for EVERY new line!
                Dim Answer53 As String, Answer54 As String
    ' a comment for EVERY new line!
                Dim Answer61 As String, Answer62 As String
    ' a comment for EVERY new line!
                Dim Answer63 As String, Answer64 As String
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                MsgBox RCPlayer1Name & " and " & RCPlayer3Name & " need to leave the room or at least turn away from the screen."
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                'player's 2 and 4 need to answer these questions first
    ' a comment for EVERY new line!
                Answer42 = InputBox("If your roommate could live anywhere, where would she live?", RCPlayer2Name & " answer this question")
    ' a comment for EVERY new line!
                Answer44 = InputBox("If your roommate could live anywhere, where would she live?", RCPlayer4Name & " answer this question")
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                Answer52 = InputBox("What color is your roommate's toothbrush?", RCPlayer2Name & " answer this question")
    ' a comment for EVERY new line!
                Answer54 = InputBox("What color is your roommate's toothbrush?", RCPlayer4Name & " answer this question")
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                Answer62 = InputBox("Which of the following does your roommate like to do least: A.Homework, B.Cleaning, C.Laundry, D.Getting out of bed?", RCPlayer2Name & " answer this question")
    ' a comment for EVERY new line!
                Answer64 = InputBox("Which of the following does your roommate like to do least: A.Homework, B.Cleaning, C.Laundry, D.Getting out of bed?", RCPlayer4Name & " answer this question")
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                MsgBox "Now " & RCPlayer1Name & " and " & RCPlayer3Name & " come back to answer the next questions!"
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                'team1 ?4 player 1 answers
    ' a comment for EVERY new line!
                Answer41 = InputBox("When asked the question: If your roommate could live anywhere, where would she live? Your roommate answered what?", RCPlayer1Name & " answer this question")
    ' a comment for EVERY new line!
                If LCase(Answer41) = LCase(Answer42) Then
    ' a comment for EVERY new line!
                MsgBox "Correct!"
    ' a comment for EVERY new line!
                team1Total = team1Total + 5
    ' a comment for EVERY new line!
                picTeam1Results.Cls
    ' a comment for EVERY new line!
                picTeam1Results.Print team1Total
    ' a comment for EVERY new line!
                Else
    ' a comment for EVERY new line!
                MsgBox "No your roommate answered " & Answer42
    ' a comment for EVERY new line!
                End If
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                'team2 ?4 player 3 answers
    ' a comment for EVERY new line!
                Answer43 = InputBox("When asked the question: If your roommate could live anywhere, where would she live? Your roommate answered what?", RCPlayer3Name & " answer this question") 'team 2 scoring
    ' a comment for EVERY new line!
                If LCase(Answer43) = LCase(Answer44) Then
    ' a comment for EVERY new line!
                MsgBox "Correct!"
    ' a comment for EVERY new line!
                team2Total = team2Total + 5
    ' a comment for EVERY new line!
                picTeam2Results.Cls
    ' a comment for EVERY new line!
                picTeam2Results.Print team2Total
    ' a comment for EVERY new line!
                Else
    ' a comment for EVERY new line!
                MsgBox "No your roommate answered " & Answer44
    ' a comment for EVERY new line!
                End If
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                'team1 ?5 player 1 answers
    ' a comment for EVERY new line!
                Answer51 = InputBox("When asked the question: What color is your roommate's toothbrush?", RCPlayer1Name & " answer this question")
    ' a comment for EVERY new line!
                If LCase(Answer51) = LCase(Answer52) Then
    ' a comment for EVERY new line!
                MsgBox "Correct!"
    ' a comment for EVERY new line!
                team1Total = team1Total + 5
    ' a comment for EVERY new line!
                picTeam1Results.Cls
    ' a comment for EVERY new line!
                picTeam1Results.Print team1Total
    ' a comment for EVERY new line!
                Else
    ' a comment for EVERY new line!
                MsgBox "No your roommate answered " & Answer52
    ' a comment for EVERY new line!
                End If
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                'team2 ?5 player 3 answers
    ' a comment for EVERY new line!
                Answer53 = InputBox("When asked the question: What color is your roommate's toothbrush? Your roommate answered what?", RCPlayer3Name & " answer this question")
    ' a comment for EVERY new line!
                If LCase(Answer53) = LCase(Answer54) Then
    ' a comment for EVERY new line!
                MsgBox "Correct!"
    ' a comment for EVERY new line!
                team2Total = team2Total + 5
    ' a comment for EVERY new line!
                picTeam2Results.Cls
    ' a comment for EVERY new line!
                picTeam2Results.Print team2Total
    ' a comment for EVERY new line!
                Else
    ' a comment for EVERY new line!
                MsgBox "No your roommate answered " & Answer54
    ' a comment for EVERY new line!
                End If
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                'team1 ?6 player 1 answers
    ' a comment for EVERY new line!
                Answer61 = InputBox("When asked the question: Which of the following does your roommate like to do least: A.Homework, B.Cleaning, C.Laundry, D.Getting out of bed? Your roommate answered what?", RCPlayer1Name & " answer this question")
    ' a comment for EVERY new line!
                If LCase(Answer61) = LCase(Answer62) Then
    ' a comment for EVERY new line!
                MsgBox "Correct!"
    ' a comment for EVERY new line!
                team1Total = team1Total + 5
    ' a comment for EVERY new line!
                picTeam1Results.Cls
    ' a comment for EVERY new line!
                picTeam1Results.Print team1Total
    ' a comment for EVERY new line!
                Else
    ' a comment for EVERY new line!
                MsgBox "No your roommate answered " & Answer62
    ' a comment for EVERY new line!
                End If
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                'team2 ?6 player 3 answers
    ' a comment for EVERY new line!
                Answer63 = InputBox("When asked the question: Which of the following does your roommate like to do least: A.Homework, B.Cleaning, C.Laundry, D.Getting out of bed? Your roommate answered what?", RCPlayer3Name & " answer this question")
    ' a comment for EVERY new line!
                If LCase(Answer63) = LCase(Answer64) Then
    ' a comment for EVERY new line!
                MsgBox "Correct!"
    ' a comment for EVERY new line!
                team2Total = team2Total + 5
    ' a comment for EVERY new line!
                picTeam2Results.Cls
    ' a comment for EVERY new line!
                picTeam2Results.Print team2Total
    ' a comment for EVERY new line!
                Else
    ' a comment for EVERY new line!
                MsgBox "No your roommate answered " & Answer64
    ' a comment for EVERY new line!
                End If
    ' a comment for EVERY new line!
                End Sub
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                Private Sub Command1_Click()
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                Open App.Path & "\Scores.txt" For Append As #1
    ' a comment for EVERY new line!
                Dim Team1Names As String
    ' a comment for EVERY new line!
                Team1Names = RCPlayer1Name + " & " + RCPlayer2Name + "          "
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                Write #1, Team1Names, team1Total
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                Close
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                Open App.Path & "\Scores.txt" For Append As #1
    ' a comment for EVERY new line!
                Dim Team2Names As String
    ' a comment for EVERY new line!
                Team2Names = RCPlayer3Name + " & " + RCPlayer4Name + "          "
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                Write #1, Team2Names, team2Total
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                Close
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                frmRoommateChallenge3.Hide
    ' a comment for EVERY new line!
                frmScores.Show
    ' a comment for EVERY new line!
                End Sub
    ' a comment for EVERY new line!

    ' a comment for EVERY new line!
                Private Sub Command2_Click()
    ' a comment for EVERY new line!
                End
    ' a comment for EVERY new line!
                End Sub
    ' a comment for EVERY new line!
