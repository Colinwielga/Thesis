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
Dim team1Total As Integer
Dim team2Total As Integer

Private Sub cmdFinalRound_Click()
Dim AnswerF11 As String, AnswerF12 As String, AnswerF13 As String, AnswerF14 As String
Dim AnswerF21 As String, AnswerF22 As String, AnswerF23 As String, AnswerF24 As String

MsgBox "This round is worth 10 points per question!"
MsgBox RCPlayer2Name & " and " & RCPlayer4Name & " please leave or look away."

AnswerF11 = InputBox("What is your roommate most afraid of?", RCPlayer1Name & " answer this question.")
AnswerF13 = InputBox("What is your roommate most afraid of?", RCPlayer3Name & " answer this question.")

MsgBox "Now " & RCPlayer2Name & " and " & RCPlayer4Name & " come back to answer the next question!"

AnswerF12 = InputBox("When asked the question: What is your roommate most afraid of? Your roommate answered what?", RCPlayer2Name & " answer this question")
If LCase(AnswerF11) = LCase(AnswerF12) Then
    MsgBox "Correct!"
    team1Total = team1Total + 10
    picTeam1Results.Cls
    picTeam1Results.Print team1Total
Else
    MsgBox "No your roommate answered " & AnswerF11
End If

AnswerF14 = InputBox("When asked the question: What is your roommate most afraid of? Your roommate answered what?", RCPlayer4Name & " answer this question")
If LCase(AnswerF14) = LCase(AnswerF13) Then
    MsgBox "Correct!"
    team2Total = team2Total + 10
    picTeam2Results.Cls
    picTeam2Results.Print team2Total
Else
    MsgBox "No your roommate answered " & AnswerF13
End If

'second question. Player 2 and 4 answer first set.  1 and 3 answer second set.
MsgBox RCPlayer1Name & " and " & RCPlayer3Name & " please leave or look away."

AnswerF22 = InputBox("What is your roommate's ideal first date: A.Romantic Dinner B.Movie C.Concert D.Long Walk on the beach?", RCPlayer2Name & " answer this question.")
AnswerF24 = InputBox("What is your roommate's ideal first date: A.Romantic Dinner B.Movie C.Concert D.Long Walk on the beach?", RCPlayer4Name & " answer this question.")

MsgBox "Now " & RCPlayer1Name & " and " & RCPlayer3Name & " come back to answer the next question!"

AnswerF21 = InputBox("When asked the question: What is your roommate's ideal first date: A.Romantic Dinner B.Movie C.Concert D.Long Walk on the beach? Your roommate answered what?", RCPlayer1Name & " answer this question")
If LCase(AnswerF21) = LCase(AnswerF22) Then
    MsgBox "Correct!"
    team1Total = team1Total + 10
    picTeam1Results.Cls
    picTeam1Results.Print team1Total
Else
    MsgBox "No your roommate answered " & AnswerF22
End If

AnswerF23 = InputBox("When asked the question: What is your roommate's ideal first date: A.Romantic Dinner B.Movie C.Concert D.Long Walk on the beach? Your roommate answered what?", RCPlayer3Name & " answer this question")
If LCase(AnswerF24) = LCase(AnswerF23) Then
    MsgBox "Correct!"
    team2Total = team2Total + 10
    picTeam2Results.Cls
    picTeam2Results.Print team2Total
Else
    MsgBox "No your roommate answered " & AnswerF24
End If

'declare a winner
If team2Total = team1Total Then
    MsgBox "It's a tie!"
ElseIf team2Total > team1Total Then
    MsgBox "Team 2 wins! Congrats " & RCPlayer3Name & " & " & RCPlayer4Name & "!", , "WINNERS"
Else
    MsgBox "Team 1 wins! Congrats " & RCPlayer1Name & " & " & RCPlayer2Name & "!", , "WINNERS"
End If

End Sub




Private Sub cmdOpinion_Click()
 frmRoommateChallenge3.Hide
 frmJudges.Show

End Sub

Private Sub cmdPlay_Click()
'load the pictures and names of the players
'player 1
picPlayer1Name.Print RCPlayer1Name
    If RCPlayer1Pic = 1 Then
            picPlayer1.Picture = LoadPicture(App.Path & "\Caitlin.bmp")
        ElseIf RCPlayer1Pic = 2 Then
            picPlayer1.Picture = LoadPicture(App.Path & "\Emily.bmp")
        ElseIf RCPlayer1Pic = 3 Then
            picPlayer1.Picture = LoadPicture(App.Path & "\me.bmp")
        ElseIf RCPlayer1Pic = 4 Then
            picPlayer1.Picture = LoadPicture(App.Path & "\heather.bmp")
        ElseIf RCPlayer1Pic = 5 Then
            picPlayer1.Picture = LoadPicture(App.Path & "\katie.bmp")
        Else: RCPlayer1Pic = 6
            picPlayer1.Picture = LoadPicture(App.Path & "\steph.bmp")
    End If
'player 2
picPlayer2Name.Print RCPlayer2Name
    If RCPlayer2Pic = 1 Then
            picPlayer2.Picture = LoadPicture(App.Path & "\Caitlin.bmp")
        ElseIf RCPlayer2Pic = 2 Then
            picPlayer2.Picture = LoadPicture(App.Path & "\Emily.bmp")
        ElseIf RCPlayer2Pic = 3 Then
            picPlayer2.Picture = LoadPicture(App.Path & "\me.bmp")
        ElseIf RCPlayer2Pic = 4 Then
            picPlayer2.Picture = LoadPicture(App.Path & "\heather.bmp")
        ElseIf RCPlayer2Pic = 5 Then
            picPlayer2.Picture = LoadPicture(App.Path & "\katie.bmp")
        Else: RCPlayer2Pic = 6
            picPlayer2.Picture = LoadPicture(App.Path & "\steph.bmp")
    End If
'player 3
picPlayer3Name.Print RCPlayer3Name
    If RCPlayer3Pic = 1 Then
            picPlayer3.Picture = LoadPicture(App.Path & "\Caitlin.bmp")
        ElseIf RCPlayer3Pic = 2 Then
            picPlayer3.Picture = LoadPicture(App.Path & "\Emily.bmp")
        ElseIf RCPlayer3Pic = 3 Then
            picPlayer3.Picture = LoadPicture(App.Path & "\me.bmp")
        ElseIf RCPlayer3Pic = 4 Then
            picPlayer3.Picture = LoadPicture(App.Path & "\heather.bmp")
        ElseIf RCPlayer3Pic = 5 Then
            picPlayer3.Picture = LoadPicture(App.Path & "\katie.bmp")
        Else: RCPlayer3Pic = 6
            picPlayer3.Picture = LoadPicture(App.Path & "\steph.bmp")
    End If
'player 4
picPlayer4Name.Print RCPlayer4Name
    If RCPlayer4Pic = 1 Then
            picPlayer4.Picture = LoadPicture(App.Path & "\Caitlin.bmp")
        ElseIf RCPlayer4Pic = 2 Then
            picPlayer4.Picture = LoadPicture(App.Path & "\Emily.bmp")
        ElseIf RCPlayer4Pic = 3 Then
            picPlayer4.Picture = LoadPicture(App.Path & "\me.bmp")
        ElseIf RCPlayer4Pic = 4 Then
            picPlayer4.Picture = LoadPicture(App.Path & "\heather.bmp")
        ElseIf RCPlayer4Pic = 5 Then
            picPlayer4.Picture = LoadPicture(App.Path & "\katie.bmp")
        Else: RCPlayer4Pic = 6
            picPlayer4.Picture = LoadPicture(App.Path & "\steph.bmp")
    End If

MsgBox RCPlayer2Name & " and " & RCPlayer4Name & " need to leave the room or at least turn away from the screen."
MsgBox "Now " & RCPlayer1Name & " and " & RCPlayer3Name & " start by clicking the Round 1 button!"


End Sub

Private Sub cmdRound1_Click()
Dim Answer11 As String, Answer12 As String
Dim Answer13 As String, Answer14 As String
Dim Answer21 As String, Answer22 As String
Dim Answer23 As String, Answer24 As String
Dim Answer31 As String, Answer32 As String
Dim Answer33 As String, Answer34 As String


team1Total = 0
team2Total = 0

Answer11 = InputBox("If your roomate had to eat one thing for the rest of her life what would it be?", RCPlayer1Name & " answer this question")
Answer13 = InputBox("If your roomate had to eat one thing for the rest of her life what would it be?", RCPlayer3Name & " answer this question")

Answer21 = InputBox("What color is your roommate's favorite shirt?", RCPlayer1Name & " answer this question")
Answer23 = InputBox("What color is your roommate's favorite shirt?", RCPlayer3Name & " answer this question")

Answer31 = InputBox("If your roomate was stranded on an island she would A. Build an incredible raft and escape; B. Panic and scream; C. Spend the time lying on the beach soaking up the sun saying 'Hey at least I'll get a killer tan'?", RCPlayer1Name & " answer this question")
Answer33 = InputBox("If your roomate was stranded on an island she would A. Build an incredible raft and escape; B. Panic and scream; C. Spend the time lying on the beach soaking up the sun saying 'Hey at least I'll get a killer tan'?", RCPlayer3Name & " answer this question")

MsgBox "Now " & RCPlayer2Name & " and " & RCPlayer4Name & " come back to answer the next questions!"

'team1 ?1
Answer12 = InputBox("When asked the question: If your roomate had to eat one thing for the rest of her life what would it be? Your roommate answered what?", RCPlayer2Name & " answer this question")
If LCase(Answer11) = LCase(Answer12) Then
    MsgBox "Correct!"
    team1Total = team1Total + 5
    picTeam1Results.Print team1Total
Else
    MsgBox "No your roommate answered " & Answer11
    picTeam1Results.Print team1Total
End If

'team2 ?2
Answer14 = InputBox("When asked the question: If your roomate had to eat one thing for the rest of her life what would it be? Your roommate answered what?", RCPlayer4Name & " answer this question") 'team 2 scoring
If LCase(Answer13) = LCase(Answer14) Then
    MsgBox "Correct!"
    team2Total = team2Total + 5
    picTeam2Results.Print team2Total
Else
    MsgBox "No your roommate answered " & Answer13
    picTeam2Results.Print team2Total
End If
    
'team1 ?2
Answer22 = InputBox("When asked the question: What color is your roommate's favorite shirt? Your roommate answered what?", RCPlayer2Name & " answer this question")
    If LCase(Answer21) = LCase(Answer22) Then
        MsgBox "Correct!"
        team1Total = team1Total + 5
        picTeam1Results.Cls
        picTeam1Results.Print team1Total
    Else
        MsgBox "No your roommate answered " & Answer21
    End If

'team2 ?2
Answer24 = InputBox("When asked the question: What color is your roommate's favorite shirt? Your roommate answered what?", RCPlayer4Name & " answer this question")
    If LCase(Answer23) = LCase(Answer24) Then
        MsgBox "Correct!"
        team2Total = team2Total + 5
        picTeam2Results.Cls
        picTeam2Results.Print team2Total
    Else
        MsgBox "No your roommate answered " & Answer23
    End If
    
'team1 ?3
Answer32 = InputBox("When asked the question: If your roomate was stranded on an island she would A. Build an incredible raft and escape; B. Panic and scream; C. Spend the time lying on the beach soaking up the sun saying 'Hey at least I'll get a killer tan'? Your roommate answered what?", RCPlayer2Name & " answer this question")
If LCase(Answer31) = LCase(Answer32) Then
    MsgBox "Correct!"
    team1Total = team1Total + 5
    picTeam1Results.Cls
    picTeam1Results.Print team1Total
Else
    MsgBox "No your roommate answered " & Answer31
End If

'team2 ?3
Answer34 = InputBox("When asked the question: If your roomate were stranded on an island she would A. Build an incredible raft and escape; B. Panic and scream; C. Spend the her time lying on the beach soaking up the sun saying 'Hey at least I'll get a killer tan'? Your roommate answered what?", RCPlayer4Name & " answer this question")
If LCase(Answer33) = LCase(Answer34) Then
    MsgBox "Correct!"
    team2Total = team2Total + 5
    picTeam2Results.Cls
    picTeam2Results.Print team2Total
Else
    MsgBox "No your roommate answered " & Answer33
End If

End Sub

Private Sub cmdRound2_Click()
Dim Answer41 As String, Answer42 As String
Dim Answer43 As String, Answer44 As String
Dim Answer51 As String, Answer52 As String
Dim Answer53 As String, Answer54 As String
Dim Answer61 As String, Answer62 As String
Dim Answer63 As String, Answer64 As String

MsgBox RCPlayer1Name & " and " & RCPlayer3Name & " need to leave the room or at least turn away from the screen."

'player's 2 and 4 need to answer these questions first
Answer42 = InputBox("If your roommate could live anywhere, where would she live?", RCPlayer2Name & " answer this question")
Answer44 = InputBox("If your roommate could live anywhere, where would she live?", RCPlayer4Name & " answer this question")

Answer52 = InputBox("What color is your roommate's toothbrush?", RCPlayer2Name & " answer this question")
Answer54 = InputBox("What color is your roommate's toothbrush?", RCPlayer4Name & " answer this question")

Answer62 = InputBox("Which of the following does your roommate like to do least: A.Homework, B.Cleaning, C.Laundry, D.Getting out of bed?", RCPlayer2Name & " answer this question")
Answer64 = InputBox("Which of the following does your roommate like to do least: A.Homework, B.Cleaning, C.Laundry, D.Getting out of bed?", RCPlayer4Name & " answer this question")

MsgBox "Now " & RCPlayer1Name & " and " & RCPlayer3Name & " come back to answer the next questions!"

'team1 ?4 player 1 answers
Answer41 = InputBox("When asked the question: If your roommate could live anywhere, where would she live? Your roommate answered what?", RCPlayer1Name & " answer this question")
If LCase(Answer41) = LCase(Answer42) Then
    MsgBox "Correct!"
    team1Total = team1Total + 5
    picTeam1Results.Cls
    picTeam1Results.Print team1Total
Else
    MsgBox "No your roommate answered " & Answer42
End If

'team2 ?4 player 3 answers
Answer43 = InputBox("When asked the question: If your roommate could live anywhere, where would she live? Your roommate answered what?", RCPlayer3Name & " answer this question") 'team 2 scoring
If LCase(Answer43) = LCase(Answer44) Then
    MsgBox "Correct!"
    team2Total = team2Total + 5
    picTeam2Results.Cls
    picTeam2Results.Print team2Total
Else
    MsgBox "No your roommate answered " & Answer44
End If
    
'team1 ?5 player 1 answers
Answer51 = InputBox("When asked the question: What color is your roommate's toothbrush?", RCPlayer1Name & " answer this question")
    If LCase(Answer51) = LCase(Answer52) Then
        MsgBox "Correct!"
        team1Total = team1Total + 5
        picTeam1Results.Cls
        picTeam1Results.Print team1Total
    Else
        MsgBox "No your roommate answered " & Answer52
    End If

'team2 ?5 player 3 answers
Answer53 = InputBox("When asked the question: What color is your roommate's toothbrush? Your roommate answered what?", RCPlayer3Name & " answer this question")
    If LCase(Answer53) = LCase(Answer54) Then
        MsgBox "Correct!"
        team2Total = team2Total + 5
        picTeam2Results.Cls
        picTeam2Results.Print team2Total
    Else
        MsgBox "No your roommate answered " & Answer54
    End If
    
'team1 ?6 player 1 answers
Answer61 = InputBox("When asked the question: Which of the following does your roommate like to do least: A.Homework, B.Cleaning, C.Laundry, D.Getting out of bed? Your roommate answered what?", RCPlayer1Name & " answer this question")
If LCase(Answer61) = LCase(Answer62) Then
    MsgBox "Correct!"
    team1Total = team1Total + 5
    picTeam1Results.Cls
    picTeam1Results.Print team1Total
Else
    MsgBox "No your roommate answered " & Answer62
End If

'team2 ?6 player 3 answers
Answer63 = InputBox("When asked the question: Which of the following does your roommate like to do least: A.Homework, B.Cleaning, C.Laundry, D.Getting out of bed? Your roommate answered what?", RCPlayer3Name & " answer this question")
If LCase(Answer63) = LCase(Answer64) Then
    MsgBox "Correct!"
    team2Total = team2Total + 5
    picTeam2Results.Cls
    picTeam2Results.Print team2Total
Else
    MsgBox "No your roommate answered " & Answer64
End If
End Sub

Private Sub Command1_Click()

Open App.Path & "\Scores.txt" For Append As #1
Dim Team1Names As String
Team1Names = RCPlayer1Name + " & " + RCPlayer2Name + "          "

Write #1, Team1Names, team1Total

Close


Open App.Path & "\Scores.txt" For Append As #1
Dim Team2Names As String
Team2Names = RCPlayer3Name + " & " + RCPlayer4Name + "          "

Write #1, Team2Names, team2Total

Close

frmRoommateChallenge3.Hide
frmScores.Show
End Sub

Private Sub Command2_Click()
 End
End Sub
