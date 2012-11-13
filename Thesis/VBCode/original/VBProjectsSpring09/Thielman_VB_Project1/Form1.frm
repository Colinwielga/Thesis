VERSION 5.00
Begin VB.Form frmpart1 
   BackColor       =   &H80000012&
   Caption         =   "Form1"
   ClientHeight    =   10485
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   10485
   ScaleWidth      =   14880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdskip 
      Caption         =   "Skip Game"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   15
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Game?"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   14
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdq4 
      Caption         =   "Question 4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   18
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7320
      TabIndex        =   13
      Top             =   7320
      Width           =   2175
   End
   Begin VB.CommandButton cmdanswer3 
      Caption         =   "Answers for # 3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5040
      TabIndex        =   12
      Top             =   8400
      Width           =   2175
   End
   Begin VB.CommandButton cmdq3 
      Caption         =   "Question 3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   18
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5040
      TabIndex        =   11
      Top             =   7320
      Width           =   2175
   End
   Begin VB.CommandButton cmdanswer2 
      Caption         =   "Answers For #2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      TabIndex        =   10
      Top             =   8400
      Width           =   2295
   End
   Begin VB.PictureBox picresultstwo 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   9
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3615
      Left            =   9840
      ScaleHeight     =   3555
      ScaleWidth      =   4875
      TabIndex        =   9
      Top             =   4800
      Width           =   4935
   End
   Begin VB.CommandButton cmdq2 
      Caption         =   "Question 2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   18
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      TabIndex        =   8
      Top             =   7320
      Width           =   2295
   End
   Begin VB.CommandButton cmdq1 
      Caption         =   "Question 1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   18
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   7
      Top             =   7320
      Width           =   2175
   End
   Begin VB.CommandButton cmdform2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Part 2"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   11160
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox lblinstructions 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Text            =   "When asked a question for each answer write out completely and in lowercase."
      Top             =   1200
      Width           =   8415
   End
   Begin VB.CommandButton cmdbegin 
      BackColor       =   &H8000000B&
      Caption         =   "Start Game"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      MaskColor       =   &H00C00000&
      TabIndex        =   4
      Top             =   1680
      Width           =   2775
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   24
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1695
      Left            =   10920
      ScaleHeight     =   1635
      ScaleWidth      =   2715
      TabIndex        =   2
      Top             =   1320
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   2160
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   4155
      ScaleWidth      =   5475
      TabIndex        =   1
      Top             =   2520
      Width           =   5535
   End
   Begin VB.Label lbltotal 
      BackColor       =   &H00000000&
      Caption         =   "Total Score"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   11280
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lbltitle 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Do You Know the Minnesota Timberwolves?"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   18
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   9615
   End
End
Attribute VB_Name = "frmpart1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Timberwolves basketball
'frmpart1
'nick thielman
'3/15
'on this form the user may play a game of trivia where their score is kept and 5 questions
'are asked. Or they may choose to skip the game and go straight to the 3rd form.
'They also may go to the 2nd form. They are asked questions which they respond to in
'inputboxes. Answers 2 and 3 are displayed in the picture box one in alphabetical order,
'the other in numerical order

Option Explicit
Dim runningtotal As Integer, position(1 To 5) As String, players(1 To 5) As String, ctr As Integer, found As Boolean
Dim benchplayers(1 To 20) As String, numbers(1 To 20) As Integer, pass As Integer, pos As Integer

'colors, pictures, font


Private Sub cmdbegin_click()
Dim score As Integer, state As String, minnesota As String
'questions pop up and score runningtotal is kept
'if statements, msgbox

MsgBox "This is a Triva Game Featuring the Minnesota Timberwolves", , "Start"

'intializes runningtotal
runningtotal = 0

state = InputBox("Let's start easy, what state do the Timberwolves play in?", "Warm Up")

'Keeps track of runningtotal of points printing it in the pic box
'If then statement
    If state = "minnesota" Then
        runningtotal = runningtotal + 1
        MsgBox "That's Correct!", , "CORRECT"
        picresults.Print runningtotal
    Else
        MsgBox "No the Timberwolves are not from " & state & ", better start trying harder", , "WRONG"
    End If
    
    
   
    
    'enable the command button for question 1
    cmdq1.Enabled = True
    
    cmdq2.Enabled = False
    cmdanswer2.Enabled = False
    cmdq3.Enabled = False
    cmdanswer3.Enabled = False
    cmdq4.Enabled = False
   
    
End Sub

Private Sub cmdq1_Click()
'Question 1
Dim KG As String

KG = InputBox("yes or no, Kevin Garrnett once played for the T-Wolves", "Question 1")

'Continues to keeps track of runningtotal of points printing it in the pic box

 If KG = "yes" Then
        picresults.Cls
        runningtotal = runningtotal + 1
        MsgBox "Your Right!", , "CORRECT"
        picresults.Print runningtotal
    Else
        MsgBox "That's Incorrect He Once Played For The T-Wolves!", , "WRONG"
    End If
    
        'disable the command button for question 1
    cmdq1.Enabled = False
    
    'enable the command button for question 2
    cmdq2.Enabled = True
    
End Sub






Private Sub cmdq2_Click()
'Question 2

Dim L As Integer, wolf As String

'asks the user to name a starter
wolf = InputBox("Can you name a current starter on the Wolves?", "Question 2")

'opens the wolves array
found = False
Open App.Path & "\wolves.txt" For Input As #1
ctr = 0
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, position(ctr), players(ctr)
Loop
Close #1

'uses a match/stop search to find the player
'continues to keep track of runningtotal of points printing it in the pic box
For L = 1 To ctr
    If wolf = players(L) Then
        found = True
        runningtotal = runningtotal + 1
    End If

 
Next L
    If (Not found) Then
        MsgBox "No,  " & wolf & ", is not a starter", , "WRONG"
    Else
        MsgBox "Good work " & wolf & " is a starter", , "CORRECT"
        picresults.Cls
        picresults.Print runningtotal
    End If
         'disable the command button for question 2
    cmdq2.Enabled = False
    
    'enable the command button for answer 2
    cmdanswer2.Enabled = True
       
    
End Sub

Private Sub cmdanswer2_Click()
'Answer #2
'alpebetically for number 2
'sorting

Dim tempplayers As String, j As Integer
Dim tempposition As String

    MsgBox "Here are the starters in alphabetical order with their positions", , "Answer"


'bubble sorts the players alphabetical in decending order
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If players(pos) > players(pos + 1) Then
            tempplayers = players(pos)
            players(pos) = players(pos + 1)
            players(pos + 1) = tempplayers
            tempposition = position(pos)
            position(pos) = position(pos + 1)
            position(pos + 1) = tempposition
            End If
    Next pos
Next pass


'prints the list of players by alphabetical order
For j = 1 To ctr
             picresultstwo.Print players(j); Tab(25); position(j)
    Next j
    
Close #1

  'disable the command button for for answer 2
    cmdanswer2.Enabled = False
    
    'enable the command button for question 3
    cmdq3.Enabled = True
    
End Sub



Private Sub cmdq3_Click()
'Question 3

Dim B As Integer, bench As String

'asks user for a bench/injured player
bench = InputBox("Can You Name A Bench or Injured Player On The T-Wolves Roster?", "Question 3")

found = False
ctr = 0

'opens array
Open App.Path & "\roster.txt" For Input As #2


Do While Not EOF(2)
    ctr = ctr + 1
    Input #2, benchplayers(ctr), numbers(ctr)
Loop
Close #2

'uses a match/stop search to find the these players
'continues to keep track of runningtotal of points right
For B = 1 To ctr
    If bench = benchplayers(B) Then
        found = True
        runningtotal = runningtotal + 1
    End If

 
Next B
    If (Not found) Then
        MsgBox "No,  " & bench & ", is not a bench player", , "WRONG"
    Else
        MsgBox "Good work " & bench & " is one, you really know the Timberwolves", , "CORRECT"
        picresults.Cls
        picresults.Print runningtotal
    End If
    
         'disable the command button for quesion 3
        cmdq3.Enabled = False
    
    
        'enable the command button for answer 3
        cmdanswer3.Enabled = True

End Sub

Private Sub cmdanswer3_Click()
Dim tempbench As String, s As Integer
Dim tempnumber As String
'Answer #3

MsgBox "Here are the bench/injured players in decending order of their jersey number", , "Answer"


'bubble sorts the players by decending order based on jersey number
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If numbers(pos) > numbers(pos + 1) Then
            tempnumber = numbers(pos)
            numbers(pos) = numbers(pos + 1)
            numbers(pos + 1) = tempnumber
            tempbench = benchplayers(pos)
            benchplayers(pos) = benchplayers(pos + 1)
            benchplayers(pos + 1) = tempbench
            End If
    Next pos
Next pass

'clears pic box
picresultstwo.Cls

'prints the players numerically
For s = 1 To ctr
             picresultstwo.Print numbers(s); Tab(25); benchplayers(s)
    Next s
    


  'disable the command button for answer3
    cmdanswer3.Enabled = False
    
    'enable the command button for question 4
    cmdq4.Enabled = True
End Sub


Private Sub cmdq4_Click()
Dim wins As Single
'Quesion 4
'select/case statement

'asks user for a win total
wins = InputBox("How Many Wins Will The Timberwolves Have Next Season? (0-82)", "Question 4")

'clears pic box
picresultstwo.Cls

'user enter how many wins they believe twolves will get and followed by a response in picresults
'continues to keep track of runningtotal of points
    Select Case wins
        Case Is = 82
            picresultstwo.Print "That's What I Think Too! You Earn A Point!"
            runningtotal = runningtotal + 1
            picresults.Cls
            picresults.Print runningtotal
        Case 74 To 81
             picresultstwo.Print "That's A Great Season!"
        Case 55 To 73
             picresultstwo.Print "That's A Good Season."
        Case 20 To 54
             picresultstwo.Print "Well...Could be better than this season"
        Case 1 To 19
             picresultstwo.Print "It must be a rebuilding year...again"
        Case Else
             picresultstwo.Print "I Don't Think That's Possbile"
            
    End Select
    
    'message box for runningtotal
    'tells the user their rank
    If runningtotal = 5 Then
        MsgBox "You are a true Timberwolves fan", , "FOUR AND THE BONUS POINT!"
    ElseIf runningtotal = 4 Then
        MsgBox "You appear to know everything about the Timberwolves", , "FOUR RIGHT!"
    ElseIf runningtotal = 3 Then
        MsgBox "You seem to know a great deal about the Timberwolves", , "THREE RIGHT!"
    ElseIf runningtotal = 2 Then
        MsgBox "You seem to be a average Timberwolf fan", , "TWO RIGHT!"
    ElseIf runningtotal = 1 Then
        MsgBox "Your must not be a Timberwolves fan", , "ONE RIGHT!"
    ElseIf runningtotal = 0 Then
        MsgBox "Were you even trying?", , "ZERO RIGHT!"
   End If
   
           'disable the command button for question 4
    cmdq4.Enabled = False
    
   
   
End Sub

Private Sub cmdform2_Click()

'closes form 1
'going to form 2
frmpart2.Show
frmpart1.Hide
End Sub



Private Sub Command1_Click()
'clears pic box 1 and 2
picresults.Cls
picresultstwo.Cls
'resets the running total
runningtotal = 0

'allows for a new game by diabling the commands accept begin
    cmdbegin.Enabled = True
    cmdq1.Enabled = False
    cmdq2.Enabled = False
    cmdanswer2.Enabled = False
    cmdq3.Enabled = False
    cmdanswer3.Enabled = False
    cmdq4.Enabled = False
   

End Sub

Private Sub cmdskip_Click()
'this skips the triva game and goes straight to the stats
'closes form 1
'going to form 3
frmpart3.Show
frmpart1.Hide
End Sub
