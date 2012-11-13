VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000012&
   Caption         =   "Form1"
   ClientHeight    =   12285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14895
   LinkTopic       =   "Form1"
   Picture         =   "Chain_Reaction1.frx":0000
   ScaleHeight     =   2.20874e7
   ScaleLeft       =   89
   ScaleMode       =   0  'User
   ScaleWidth      =   48504.23
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command16 
      Caption         =   "Start Round 1"
      Height          =   1695
      Left            =   600
      TabIndex        =   23
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Final Bonus Round"
      Height          =   1695
      Left            =   13320
      TabIndex        =   22
      Top             =   10680
      Width           =   2175
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Guess for Final Bonus Round"
      Height          =   1695
      Left            =   15600
      TabIndex        =   21
      Top             =   10680
      Width           =   2055
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Guess for Bonus Round 2"
      Height          =   1335
      Left            =   15720
      TabIndex        =   20
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Bonus Round 2"
      Height          =   1335
      Left            =   13920
      TabIndex        =   19
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H8000000D&
      Caption         =   "Start Round 3"
      Height          =   1335
      Left            =   600
      TabIndex        =   18
      Top             =   11160
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Guess for Bonus Round 1"
      Height          =   1335
      Left            =   15720
      TabIndex        =   17
      Top             =   5040
      Width           =   1695
   End
   Begin VB.PictureBox Pic6 
      Height          =   495
      Left            =   6600
      ScaleHeight     =   435
      ScaleWidth      =   5715
      TabIndex        =   16
      Top             =   3840
      Width           =   5775
   End
   Begin VB.PictureBox Pic5 
      Height          =   495
      Left            =   6600
      ScaleHeight     =   435
      ScaleWidth      =   5715
      TabIndex        =   15
      Top             =   3240
      Width           =   5775
   End
   Begin VB.PictureBox Pic4 
      Height          =   495
      Left            =   6600
      ScaleHeight     =   435
      ScaleWidth      =   5715
      TabIndex        =   14
      Top             =   2640
      Width           =   5775
   End
   Begin VB.PictureBox Pic3 
      Height          =   495
      Left            =   6600
      ScaleHeight     =   435
      ScaleWidth      =   5715
      TabIndex        =   13
      Top             =   2040
      Width           =   5775
   End
   Begin VB.PictureBox Pic2 
      Height          =   495
      Left            =   6600
      ScaleHeight     =   435
      ScaleWidth      =   5715
      TabIndex        =   12
      Top             =   1440
      Width           =   5775
   End
   Begin VB.PictureBox Pic1 
      Height          =   495
      Left            =   6600
      ScaleHeight     =   435
      ScaleWidth      =   5715
      TabIndex        =   11
      Top             =   840
      Width           =   5775
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Rules"
      Height          =   1935
      Left            =   4800
      TabIndex        =   10
      Top             =   10440
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      Height          =   1935
      Left            =   9240
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   10440
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Start Round 2"
      Height          =   1335
      Left            =   600
      TabIndex        =   8
      Top             =   9600
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Bonus Round 1"
      Height          =   1335
      Left            =   13920
      TabIndex        =   7
      Top             =   5040
      Width           =   1695
   End
   Begin VB.PictureBox PicScore2 
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   1755
      TabIndex        =   6
      Top             =   2280
      Width           =   1815
   End
   Begin VB.PictureBox PicScore1 
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   1755
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Team 2"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Team 1"
      Height          =   495
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Letter Above"
      Height          =   1215
      Left            =   4800
      TabIndex        =   2
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Letter Below"
      Height          =   1095
      Left            =   4800
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Start Chain Reaction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Double click to Play Me"
      Height          =   375
      Left            =   7440
      TabIndex        =   25
      Top             =   9480
      Width           =   1095
   End
   Begin VB.OLE OLE1 
      Class           =   "Package"
      Height          =   855
      Left            =   7320
      OleObjectBlob   =   "Chain_Reaction1.frx":6F75
      SourceDoc       =   "C:\Documents and Settings\m1hailesela\Desktop\Cs130\vb PROJECT\j_game.mp3"
      TabIndex        =   24
      Top             =   9000
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Chain Reaction
'Main Screen
'Moguss Haile-Selaissie and Charlie Russell
'Written March 18th, 2009
'The purpose of this form is to open the game "Chain Reaction" for two opponents.
'The form will display the first and last words of the array and the opponents will
'in turn guess the remaining words of the array.
'Opponents will earn different amounts of money for correct answers.
'By completing the chain, opponents will then have a chance to gain more money from bonus rounds.


Dim first(1 To 500) As String, complete As Integer, hidden As Integer
Public Ctr As Integer, A As Integer, B As Integer, I As Integer, K As Integer, L As Integer, M As Integer
Public team As Boolean, score1 As Integer, score2 As Integer
Dim Bonus(1 To 20) As String
Public scoreB As Integer, scoreB2 As Integer, Ctr2 As Integer
Dim G As Integer, H As Integer, ctr3 As Integer
Dim Bonus2(1 To 20) As String, Ctr4 As Integer, Ctr5 As Integer
Dim BonusF(1 To 20) As String, Ctr6 As Integer



Private Sub Command1_Click()
'this comand introduces the game to the user and prompts he or she to start
'by reading the rules for the details of the game.


MsgBox "Welcome to Chain Reaction!", , "Welcome"
MsgBox " Click the (RULES button) for Instructions on how to play", , " Click Rules"
 Command1.Enabled = False
 Command2.Enabled = False
 Command3.Enabled = False
 Command6.Enabled = False
 Command7.Enabled = False
 Command10.Enabled = False
 Command11.Enabled = False
 Command12.Enabled = False
 Command13.Enabled = False
 Command14.Enabled = False
 Command15.Enabled = False
 Command16.Enabled = False
 
End Sub





Private Sub Command16_Click()
'This command button (Start Round 1) reads the file into an array
'It displays the first and last words, in which the user will attempt to
'connect the chain
A = 2
B = 5
complete = 0
team = True
Open App.Path & "\List1.txt" For Input As #1
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, first(Ctr)
    Loop
Pic1.Print first(1)
Pic6.Print first(Ctr)
Close #1

Command16.Enabled = False
Command3.Enabled = True
Command2.Enabled = True

End Sub

'Button for rules to appear. Once User uses the command they are
'allowed to used to commands for inputing letters

Private Sub Command9_Click()
Form1.Hide
Form2.Show
Command16.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command6.Enabled = False
Command9.Enabled = False
Form2.Command2.Enabled = False



End Sub

'Command button to incrementally output the first letter of the first
'Word directly under the displayed word.  Every time the button is clicked
'The next letter in that word with appear.
Private Sub Command2_Click()

'variable for the inputbox guess
Dim guess As String


Select Case A
'to only select the specific word in the array
Case 2
' counter to keep track of which letter to display
I = I + 1
    Select Case I
    'where and how each letter is displayed
        Case 1
            Pic2.Print Left(first(A), I)
        Case 2
            Pic2.Cls
            Pic2.Print Left(first(A), I)
        Case 3
            Pic2.Cls
            Pic2.Print Left(first(A), I)
        Case 4
             Pic2.Cls
            Pic2.Print Left(first(A), I)
        Case 5
            Pic2.Cls
             Pic2.Print Left(first(A), I)
        Case 6
            Pic2.Cls
            Pic2.Print Left(first(A), I)
        

End Select
'The user will then recieve a
'msgbox informing of the new display and an
'inputbox to submit a guess.
MsgBox "The first letter(s) under " & first(A - 1) & " is " & Left(first(A), I), , "Letter Below"
guess = InputBox("Guess the word starting with " & Left(first(A), I), "Take a Guess")
            
           
        If guess <> first(A) Then
            MsgBox "That is the wrong answer! Next teams turn", , "Wrong"
                'if statement to switch the teams automatically
                'Based on if the guess is wrong.
                If team = True Then
                    team = False
                Else
                'team 1 stays true
                    team = True
                End If
            
        Else
        'If the guess is correct
            MsgBox "That is Correct! You just earned 100 Dollars!", , "Nice Work"
        'adds and displays 100 points into picture box based on which team
        'is guessing correctly.
            Pic2.Cls
            Pic2.Print first(A)
            
                'when team 1 is guessing the boolean express is equal to true
                ' When false it is equal to team 2
                If team = True Then
                    score1 = score1 + 100
                    PicScore1.Cls
                   PicScore1.Print FormatCurrency(score1)
                    PicScore2.Cls
                    PicScore2.Print FormatCurrency(score2)
                Else
                    score2 = score2 + 100
                    PicScore2.Cls
                    PicScore2.Print FormatCurrency(score2)
                    PicScore1.Cls
                    PicScore1.Print FormatCurrency(score1)
                End If
            'once word is guessed right adding 1 to the case A to move
            'to the next word in the array
            A = A + 1
            
            'complete is the variable to keep track of when the array words are all guessed correctly
            'once all words are found the counter then stop the round.
            complete = complete + 1
        End If
        
               
         
Case 3
K = K + 1
    Select Case K
        Case 1
            Pic3.Cls
            Pic3.Print Left(first(A), K)
        Case 2
            Pic3.Cls
            Pic3.Print Left(first(A), K)
        Case 3
            Pic3.Cls
            Pic3.Print Left(first(A), K)
        Case 4
            Pic3.Cls
            Pic3.Print Left(first(A), K)
        Case 5
            Pic3.Cls
            Pic3.Print Left(first(A), K)
        Case 6
            Pic3.Cls
            Pic3.Print Left(first(A), K)
    
End Select
' inform the user of the letter that is displayed and to start with the Capital letter
MsgBox "The first letter(s) under " & first(A - 1) & " is " & Left(first(A), K), , "Letter Below"
' inputbox appears to let the user guess the word
guess = InputBox("Guess the word starting with " & Left(first(A), K), "Take a Guess")
     ' searches the specific line in the array to match with the guess
     'if incorrect it displays a message box
     If guess <> first(A) Then
            MsgBox "That is the wrong answer! Next teams turn", , "Sorry"
            'if statement to change teams after a wrong guess
            If team = True Then
                team = False
            Else
                team = True
            End If
            
        Else
            'correct guesses, earn dollars.
            MsgBox "That is Correct! You just earned 100 Dollars!", , "You Rock!"
            'clears the individual letters and then prints the whole word
            Pic3.Cls
            Pic3.Print first(A)
            ' print the dollar figure is picture boxes for the correct team
                If team = True Then
                    score1 = score1 + 100
                    PicScore1.Cls
                    PicScore1.Print FormatCurrency(score1)
                    PicScore2.Cls
                    PicScore2.Print FormatCurrency(score2)
                Else
                    score2 = score2 + 100
                    PicScore2.Cls
                    PicScore2.Print FormatCurrency(score2)
                    PicScore1.Cls
                    PicScore1.Print FormatCurrency(score1)
                End If
            A = A + 1
            complete = complete + 1
        End If
         
'explained earlier
Case 4
L = L + 1
    Select Case L
        Case 1
            Pic4.Print Left(first(A), L)
        Case 2
            Pic4.Cls
            Pic4.Print Left(first(A), L)
        Case 3
            Pic4.Cls
            Pic4.Print Left(first(A), L)
        Case 4
            Pic4.Cls
            Pic4.Print Left(first(A), L)
        Case 5
            Pic5.Cls
            Pic5.Print Left(first(A), L)
        Case 6
            Pic6.Cls
            Pic6.Print Left(first(A), L)
    
End Select
MsgBox "The first letter(s) under " & first(A - 1) & " is " & Left(first(A), L), , "Letter Below"
guess = InputBox("Guess the word starting with " & Left(first(A), L), " Take a Guess")
    
       If guess <> first(A) Then
            MsgBox "That is the wrong answer! Next teams turn", , "Wrong"
            If team = True Then
                team = False
            Else
                team = True
            End If
            
        Else
            MsgBox "That is Correct! You just earned 100 Dollars!", , "Right on!"
            Pic4.Cls
            Pic4.Print first(A)
            complete = complete + 1
                If team = True Then
                    score1 = score1 + 100
                    PicScore1.Cls
                    PicScore1.Print FormatCurrency(score1)
                    PicScore2.Cls
                    PicScore2.Print FormatCurrency(score2)
                Else
                    score2 = score2 + 100
                    PicScore2.Cls
                    PicScore2.Print FormatCurrency(score1)
                    PicScore1.Cls
                    PicScore1.Print FormatCurrency(score2)
                End If
            A = A + 1
            
        End If

         
Case 5
M = M + 1
    Select Case M
        Case 1
            Pic5.Print Left(first(A), M)
        Case 2
            Pic5.Cls
            Pic5.Print Left(first(A), M)
        Case 3
            Pic5.Cls
            Pic5.Print Left(first(A), M)
        Case 4
            Pic5.Cls
            Pic5.Print Left(first(A), M)
        Case 5
            Pic5.Cls
            Pic5.Print Left(first(A), M)
        Case 6
            Pic5.Cls
            Pic5.Print Left(first(A), M)
    
End Select
MsgBox "The first letter(s) under " & first(A - 1) & " is " & Left(first(A), M), , "Letter Below"
guess = InputBox("Guess the word starting with " & Left(first(A), M), "Take A Guess")
    
       If guess <> first(A) Then
            MsgBox "That is the wrong answer! Next teams turn", , "Wrong"
            If team = True Then
                team = False
            Else
                team = True
            End If
            
        Else
            MsgBox "That is Correct! You just earned 100 Dollars!", , "Correct"
            Pic5.Cls
            Pic5.Print first(A)
            complete = complete + 1
                If team = True Then
                    score1 = score1 + 100
                    PicScore1.Cls
                    PicScore1.Print FormatCurrency(score1)
                    PicScore2.Cls
                    PicScore2.Print FormatCurrency(score2)
                Else
                    score2 = score2 + 100
                    PicScore2.Cls
                    PicScore2.Print FormatCurrency(score1)
                    PicScore1.Cls
                    PicScore1.Print FormatCurrency(score2)
                End If
            A = A + 1
             
        End If
End Select
' Once all the words have been found by the users.  The function below will end the round and prompt
' the team who completed the last word that he or she has qualified for the bonus round

If complete = 4 Then
    MsgBox "End of Round .", , "Round over"
    
        
        ' New "if" statement to distinguish which button showed be enabled
        ' based on each round. And the message box to be out putted.
        
            If hidden < 1 Then
                hidden = hidden + 1
                Command6.Enabled = True
                Command2.Enabled = False
                Command3.Enabled = False
                    If team = True Then
                        MsgBox "Team 1 Qualifies for the bonus round.", , "Congrats team 1"
                        MsgBox "Click the button (Bonus round )-To start", , "Congrats team 1"
                    ElseIf team = False Then
                        MsgBox "Team 2 Qualifies for the bonus round.", , "Congrats team 2"
                        MsgBox "Click the button (Bonus round )-To start", , "Congrats team 2"
                    End If
            ElseIf hidden = 1 Then
                hidden = hidden + 1
                Command12.Enabled = True
                Command2.Enabled = False
                Command3.Enabled = False
                    If team = True Then
                        MsgBox "Team 1 Qualifies for the bonus round.", , "Congrats team 1"
                        MsgBox "Click the button (Bonus round )-To start", , "Congrats team 1"
                    ElseIf team = False Then
                        MsgBox "Team 2 Qualifies for the bonus round.", , "Congrats team 2"
                        MsgBox "Click the button (Bonus round )-To start", , "Congrats team 2"
                    End If
            ElseIf hidden = 2 Then
                Command15.Enabled = True
                Command2.Enabled = False
                Command3.Enabled = False
                    If score1 > score2 Then
                        MsgBox "Team 1 is the winner for today!", , "Congrats team 1"
                        MsgBox "Click the button (Final Bonus Round) for your chance to win the Grand Prize.", , "Winner."
                    Else
                       MsgBox "Team 2 is the winner for today!", , "Congrats team 2"
                       MsgBox "Click the button (Final Bonus Round) for your chance to win the Grand Prize.", , "Winner."
                    End If

End If
End If
End Sub


Private Sub Command3_Click()
Dim guess As String

'This command is the same as the previous command. Except for
'It is working in reverse order.  The variable are the same because
'It any case the words could overlap.  This allow no words to be guessed
'more then once
Select Case B

Case 5
M = M + 1
    Select Case M
        Case 1
            Pic5.Print Left(first(B), M)
        Case 2
            Pic5.Cls
            Pic5.Print Left(first(B), M)
        Case 3
            Pic5.Cls
            Pic5.Print Left(first(B), M)
        Case 4
             Pic5.Cls
            Pic5.Print Left(first(B), M)
        Case 5
            Pic5.Cls
            Pic5.Print Left(first(B), M)
        Case 6
            Pic5.Cls
            Pic5.Print Left(first(B), M)
    

    End Select
MsgBox "The first letter(s) Above " & first(B + 1) & " is " & Left(first(B), M), , "Letter Above"
guess = InputBox("Guess the word starting with " & Left(first(B), M), "Take A Guess")
            
           
        If guess <> first(B) Then
            MsgBox "That is the wrong answer! Next teams turn", , "Wrong"
                If team = True Then
                    team = False
                Else
                    team = True
                End If
            
        Else
            MsgBox "That is Correct! You just earned 100 Dollars!", , "Correct"
            Pic5.Cls
            Pic5.Print first(B)
            complete = complete + 1
                If team = True Then
                    score1 = score1 + 100
                    PicScore1.Cls
                    PicScore1.Print FormatCurrency(score1)
                    PicScore2.Cls
                    PicScore2.Print FormatCurrency(score2)
                Else
                    score2 = score2 + 100
                    PicScore2.Cls
                    PicScore2.Print FormatCurrency(score1)
                    PicScore1.Cls
                    PicScore1.Print FormatCurrency(score2)
                End If
            B = B - 1
            
        End If
        
               
         
Case 4
L = L + 1
    Select Case L
        Case 1
            Pic4.Cls
            Pic4.Print Left(first(B), L)
        Case 2
             Pic4.Cls
             Pic4.Print Left(first(B), L)
        Case 3
             Pic4.Cls
            Pic4.Print Left(first(B), L)
        Case 4
             Pic4.Cls
            Pic4.Print Left(first(B), L)
        Case 5
            Pic4.Cls
            Pic4.Print Left(first(B), L)
        Case 6
            Pic4.Cls
            Pic4.Print Left(first(B), L)
            
            
    End Select

MsgBox "The first letter(s) Above " & first(B + 1) & " is " & Left(first(B), L), , "Letter Above"

guess = InputBox("Guess the word starting with " & Left(first(B), L), "Take A Guess")
        
     If guess <> first(B) Then
            MsgBox "That is the wrong answer! Next teams turn", , "Wrong"
            If team = True Then
                team = False
            Else
                team = True
            End If
            
        Else
            MsgBox "That is Correct! You just earned 100 points!", , "Correct"
            Pic4.Cls
            Pic4.Print first(B)
            complete = complete + 1
                If team = True Then
                    score1 = score1 + 100
                    PicScore1.Cls
                    PicScore1.Print FormatCurrency(score1)
                    PicScore2.Cls
                    PicScore2.Print FormatCurrency(score2)
                Else
                    score2 = score2 + 100
                    PicScore2.Cls
                    PicScore2.Print FormatCurrency(score1)
                    PicScore1.Cls
                    PicScore1.Print FormatCurrency(score2)
                End If
            B = B - 1
            
        End If
         
        
Case 3
K = K + 1
    Select Case K
        Case 1
            Pic3.Print Left(first(B), K)
        Case 2
            Pic3.Cls
            Pic3.Print Left(first(B), K)
        Case 3
            Pic3.Cls
            Pic3.Print Left(first(B), K)
        Case 4
            Pic3.Cls
            Pic3.Print Left(first(B), K)
        Case 5
            Pic3.Cls
            Pic3.Print Left(first(B), K)
        Case 6
            Pic3.Cls
            Pic3.Print Left(first(B), K)
    
    End Select
MsgBox "The first letter(s) Above " & first(B + 1) & " is " & Left(first(B), K), , "Letter Above"
guess = InputBox("Guess the word starting with " & Left(first(B), K), "Take A Guess")
    
       If guess <> first(B) Then
            MsgBox "That is the wrong answer! Next teams turn", , "Wrong"
            If team = True Then
                team = False
            Else
                team = True
            End If
            
        Else
            MsgBox "That is Correct! You just earned 100 Dollars!", , "Correct"
            Pic3.Cls
            Pic3.Print first(B)
            complete = complete + 1
                If team = True Then
                    score1 = score1 + 100
                    PicScore1.Cls
                    PicScore1.Print FormatCurrency(score1)
                    PicScore2.Cls
                    PicScore2.Print FormatCurrency(score2)
                Else
                    score2 = score2 + 100
                    PicScore2.Cls
                    PicScore2.Print FormatCurrency(score1)
                    PicScore1.Cls
                    PicScore1.Print FormatCurrency(score2)
                End If
            B = B - 1
        
        End If

         
Case 2
I = I + 1
    Select Case I
        Case 1
            Pic2.Print Left(first(B), I)
        Case 2
            Pic2.Cls
            Pic2.Print Left(first(B), I)
        Case 3
            Pic2.Cls
            Pic2.Print Left(first(B), I)
        Case 4
            Pic2.Cls
            Pic2.Print Left(first(B), I)
        Case 5
            Pic2.Cls
            Pic2.Print Left(first(B), I)
        Case 6
            Pic2.Cls
            Pic2.Print Left(first(B), I)
    
    End Select
MsgBox "The first letter(s) Above " & first(B + 1) & " is " & Left(first(B), I), , "Letter Above"
guess = InputBox("Guess the word starting with " & Left(first(B), I), "Take A Guess")
    
       If guess <> first(B) Then
            MsgBox "That is the wrong answer! Next teams turn", , "Wrong"
            If team = True Then
                team = False
            Else
                team = True
            End If
            
        Else
            MsgBox "That is Correct! You just earned 100 Dollars!", , "Correct"
            Pic2.Cls
            Pic2.Print first(B)
            complete = complete + 1
                If team = True Then
                    score1 = score1 + 100
                    PicScore1.Cls
                    PicScore1.Print FormatCurrency(score1)
                    PicScore2.Cls
                    PicScore2.Print FormatCurrency(score2)
                Else
                    score2 = score2 + 100
                    PicScore2.Cls
                    PicScore2.Print FormatCurrency(score1)
                    PicScore1.Cls
                    PicScore1.Print FormatCurrency(score2)
                End If
            B = B - 1
          
        End If
' Once all the words have been found by the users.  The function below will end the round and prompt
' the team who completed the last word that he or she has qualified for the bonus round
If complete = 4 Then
    MsgBox "End of Round .", , "Round over"
    
        
        ' New "if" statement to distinguish which button showed be enabled
        ' based on each round. And the message box to be out putted.
        
            If hidden < 1 Then
                hidden = hidden + 1
                Command6.Enabled = True
                Command2.Enabled = False
                Command3.Enabled = False
                    If team = True Then
                        MsgBox "Team 1 Qualifies for the bonus round.", , "Congrats team 1"
                        MsgBox "Click the button (Bonus round )-To start", , "Congrats team 1"
                    ElseIf team = False Then
                        MsgBox "Team 2 Qualifies for the bonus round.", , "Congrats team 2"
                        MsgBox "Click the button (Bonus round )-To start", , "Congrats team 2"
                    End If
            ElseIf hidden = 1 Then
                hidden = hidden + 1
                Command12.Enabled = True
                Command2.Enabled = False
                Command3.Enabled = False
                    If team = True Then
                        MsgBox "Team 1 Qualifies for the bonus round.", , "Congrats team 1"
                        MsgBox "Click the button (Bonus round )-To start", , "Congrats team 1"
                    ElseIf team = False Then
                        MsgBox "Team 2 Qualifies for the bonus round.", , "Congrats team 2"
                        MsgBox "Click the button (Bonus round )-To start", , "Congrats team 2"
                    End If
            ElseIf hidden = 2 Then
                Command15.Enabled = True
                Command2.Enabled = False
                Command3.Enabled = False
                    If score1 > score2 Then
                        MsgBox "Team 1 is the winner for today!", , "Congrats team 1"
                        MsgBox "Click the button (Final Bonus Round) for your chance to win the Grand Prize.", , "Winner."
                    Else
                       MsgBox "Team 2 is the winner for today!", , "Congrats team 2"
                       MsgBox "Click the button (Final Bonus Round) for your chance to win the Grand Prize.", , "Winner."
                    End If
End If
End If
End Select
End Sub


'this command starts bonus round 1
Private Sub Command6_Click()
'prevents user from clicking buttons. Attempting to gain letters.
Command3.Enabled = False
Command2.Enabled = False
Command1.Enabled = False
Command10.Enabled = True
Command6.Enabled = False


'Clear all picture boxes.
Pic1.Cls
Pic2.Cls
Pic3.Cls
Pic4.Cls
Pic5.Cls
Pic6.Cls
'open file into an array
'for the user to guess at.

Open App.Path & "\Bonus1.txt" For Input As #1
    Do While Not EOF(1)
        Ctr2 = Ctr2 + 1
        Input #1, Bonus(Ctr2)
    Loop
'Displays the first and last words
'But the first letters of the middle two words are displayed
Pic2.Print Bonus(1)
Pic3.Print Left(Bonus(2), 1)
Pic4.Print Left(Bonus(3), 1)
Pic5.Print Bonus(Ctr2)
Close #1

'inform user of How the bonus round works
MsgBox "Complete the chain by clicking, (GUESS FOR BONUS ROUND 1). No additional letters will be given.", , "Bonus Round 1"
MsgBox "300 Bonus Dollars are at stake. Good luck!", , "300 Dollars"

End Sub

'By clicking this command the user is ready to
'attempt a guess at the two middle words

Private Sub Command10_Click()
Dim guessB As String, guessB2 As String
'inputbox for the two seperate guesses.
guessB = InputBox("What is your guess for the word starting with " & Left(Bonus(2), 1) & " ?", "Take A Guess")
guessB2 = InputBox("What is your guess for the word starting with " & Left(Bonus(3), 1) & " ?", "Take A Guess")

Command7.Enabled = True
Command6.Enabled = False
Command10.Enabled = False


'Prints the user's guess
Pic3.Cls
Pic3.Print guessB
Pic4.Cls
Pic4.Print guessB2

'If simply one of the guesses are wrong the user gets no points
If guessB <> Bonus(2) Then
    Pic3.Cls
    Pic3.Print Bonus(2)
    Pic4.Cls
    Pic4.Print Bonus(3)
    MsgBox " Sorry! That is incorrect.", , "Sorry"
     MsgBox "The answer is " & Bonus(2) & " and " & Bonus(3), , "Answer"
    MsgBox " Ready for Round 2? --Start by clicking the (Round 2) button", , "Round 2"
'If simply one of the guesses are wrong the user gets no points
ElseIf guessB2 <> Bonus(3) Then
           'prints correct answers
            Pic3.Cls
            Pic3.Print Bonus(2)
            Pic4.Cls
            Pic4.Print Bonus(3)
     MsgBox " Sorry! That is incorrect.", , "Sorry"
      MsgBox "The answer is " & Bonus(2) & " and " & Bonus(3), , "Answer"
     MsgBox " Ready for Round 2? --Start by clicking the (Round 2) button", , "Round 2"
Else
'If both guesses are correct
    MsgBox " Congrats! Your team just earned another 300 points!", , "Correct"
    MsgBox " Ready for Round 2? --Start by clicking the (Round 3) button", , "Round 2"
    If team = True Then
                    score1 = score1 + 300
                   
                Else
                    score2 = score2 + 300
        
                End If
End If
'prints scores after bonus round is complete
PicScore1.Cls
PicScore1.Print FormatCurrency(score1)
PicScore2.Cls
PicScore2.Print FormatCurrency(score2)



End Sub
'Starts round 2
Private Sub Command7_Click()
'prevents user to go back to round 1
'Allows the guessing buttons to be accessible
Command3.Enabled = True
Command2.Enabled = True
Command1.Enabled = False
Command7.Enabled = False
Command12.Enabled = False


'Clear all picutre boxes
Pic1.Cls
Pic2.Cls
Pic3.Cls
Pic4.Cls
Pic5.Cls
Pic6.Cls
'Reset cases variables to initial amount
A = 2
B = 5
complete = 0


 'opens round 2 list in to array
 Open App.Path & "\List2.txt" For Input As #1
    Do While Not EOF(1)
        ctr3 = ctr3 + 1
        Input #1, first(ctr3)
    Loop
'Prints the first and last words of the array
Pic1.Print first(1)
Pic6.Print first(ctr3)
Close #1
'Set variables to zero
I = 0
K = 0
L = 0
M = 0

End Sub

'Command button for bonus round2
Private Sub Command12_Click()
Command12.Enabled = False
Command13.Enabled = True
Command2.Enabled = False
Command3.Enabled = False

Pic1.Cls
Pic2.Cls
Pic3.Cls
Pic4.Cls
Pic5.Cls
Pic6.Cls

Open App.Path & "\Bonus2.txt" For Input As #1
    Do While Not EOF(1)
        Ctr4 = Ctr4 + 1
        Input #1, Bonus2(Ctr4)
    Loop
Pic2.Print Bonus2(1)
Pic3.Print Left(Bonus2(2), 1)
Pic4.Print Left(Bonus2(3), 1)
Pic5.Print Bonus2(Ctr4)
Close #1
MsgBox "Complete the chain by clicking, (GUESS FOR BONUS ROUND 2). No additional letters will be given.", , "Bonus Round 2"
MsgBox "500 Bonus Dollars are at stake. Good luck!", , "$500"

End Sub
'Commmand buttong to display an input box for a guess in the 2nd bonus round
Private Sub Command13_Click()
Dim guessB As String, guessB2 As String
Command11.Enabled = True
Command13.Enabled = False

guessB = InputBox("What is your guess for the word starting with " & Left(Bonus2(2), 1) & " ?", "Take A Guess")
guessB2 = InputBox("What is your guess for the word starting with " & Left(Bonus2(3), 1) & " ?", "Take A Guess")

Pic3.Cls
Pic3.Print guessB
Pic4.Cls
Pic4.Print guessB2

If guessB <> Bonus2(2) Then
    Pic3.Cls
    Pic3.Print Bonus2(2)
    Pic4.Cls
    Pic4.Print Bonus2(3)
    MsgBox " Sorry! That is incorrect.", , "Sorry"
    MsgBox "The answer is " & Bonus2(2) & " and " & Bonus2(3), , "Answer"
    MsgBox " Ready for Round 3? --Start by clicking the (Round 3) button", , "Round 3"
ElseIf guessB2 <> Bonus2(3) Then
            Pic3.Cls
            Pic3.Print Bonus2(2)
            Pic4.Cls
            Pic4.Print Bonus2(3)
     MsgBox " Sorry! That is incorrect.", , "Sorry"
      MsgBox "The answer is " & Bonus2(2) & " and " & Bonus2(3), , "Answer"
     MsgBox " Ready for Round 3? --Start by clicking the (Round 3) button", , "Round 3"
Else
    MsgBox " Congrats! Your team just earned another 300 points!", , "Correct"
    MsgBox " Ready for Round 3? --Start by clicking the (Round 3) button", , "Round 3"
    If team = True Then
                   score1 = score1 + 300
                   
                Else
                    score2 = score2 + 300
        
                End If
End If

PicScore1.Cls
PicScore1.Print FormatCurrency(score1)
PicScore2.Cls
PicScore2.Print FormatCurrency(score2)
End Sub
'Start round 3
Private Sub Command11_Click()
Command11.Enabled = False
Command2.Enabled = True
Command3.Enabled = True
Command15.Enabled = True

Pic1.Cls
Pic2.Cls
Pic3.Cls
Pic4.Cls
Pic5.Cls
Pic6.Cls

 A = 2
 B = 5
 complete = 0
 
'open file three in to an array
 Open App.Path & "\List3.txt" For Input As #1
    Do While Not EOF(1)
        Ctr5 = Ctr5 + 1
        Input #1, first(Ctr5)
    Loop
Pic1.Print first(1)
Pic6.Print first(Ctr5)
Close #1
I = 0
K = 0
L = 0
M = 0


End Sub
'Command button for the final Bonus round
Private Sub Command15_Click()
Command15.Enabled = False
Command14.Enabled = True
Command2.Enabled = False
Command3.Enabled = False


Pic1.Cls
Pic2.Cls
Pic3.Cls
Pic4.Cls
Pic5.Cls
Pic6.Cls

Open App.Path & "\bonusfinal.txt" For Input As #1
    Do While Not EOF(1)
        Ctr6 = Ctr6 + 1
        Input #1, BonusF(Ctr6)
    Loop
Pic1.Print BonusF(1)
Pic2.Print Left(BonusF(2), 1)
Pic3.Print Left(BonusF(3), 1)
Pic4.Print Left(BonusF(4), 1)
Pic5.Print BonusF(Ctr6)
Close #1
MsgBox "Complete the chain by clicking the button (GUESS FOR FINAL BONUS ROUND). No additional letters will be given.", , "Good Luck"
MsgBox "$$$25,000$$$ DOLLARS are at stake. Good luck!", , "Good Luck"
End Sub


'Command button to display an input box for the user to guess the answers in the final bonus round.

Private Sub Command14_Click()
Dim guessB As String, guessB2 As String, guessB3 As String

guessB = InputBox("What is your guess for the word starting with " & Left(BonusF(2), 1) & " ?", "Guess")
guessB2 = InputBox("What is your guess for the word starting with " & Left(BonusF(3), 1) & " ?", "Guess")
guessB3 = InputBox("What is your guess for the word starting with " & Left(BonusF(4), 1) & " ?", "Guess")

Pic2.Cls
Pic2.Print guessB
Pic3.Cls
Pic3.Print guessB2
Pic4.Cls
Pic4.Print guessB3

If guessB <> BonusF(2) Then
    Pic2.Cls
    Pic2.Print BonusF(2)
    Pic3.Cls
    Pic3.Print BonusF(3)
    Pic4.Cls
    Pic4.Print BonusF(4)
    MsgBox " Sorry! That is incorrect.", , "Sorry"
    MsgBox "The answer is " & BonusF(2) & ", " & BonusF(3) & ", " & " and " & BonusF(4), , "Answer"
    MsgBox " You get no money, but Thanks for Playing!", , "Sorry"
ElseIf guessB2 <> BonusF(3) Then
            Pic2.Cls
            Pic2.Print BonusF(2)
            Pic3.Cls
            Pic3.Print BonusF(3)
            Pic4.Cls
            Pic4.Print BonusF(4)
     MsgBox " Sorry! That is incorrect.", , "Sorry"
     MsgBox "The answer is " & BonusF(2) & ", " & BonusF(3) & ", " & " and " & BonusF(4), , "Answer"
     MsgBox " You get no money, but Thanks for Playing!", , "Sorry"
ElseIf guessB3 <> BonusF(4) Then
    Pic2.Cls
            Pic2.Print BonusF(2)
            Pic3.Cls
            Pic3.Print BonusF(3)
            Pic4.Cls
            Pic4.Print BonusF(4)
    MsgBox " Sorry! That is incorrect.", , "Sorry"
    MsgBox "The answer is " & BonusF(2) & ", " & BonusF(3) & ", " & " and " & BonusF(4), , "Answer"
    MsgBox " You get no money, but Thanks for Playing!", , "Sorry"
     If team = True Then
                    score1 = 0
                   
                Else
                    score2 = 0
        
                End If
Else
    MsgBox " Congrats! You have just won ***************$25,000 DOLLARS!***************", , "Winner!"
    MsgBox " Thanks for Playing! Play again soon!", , "End of Game"
    If team = True Then
                    score1 = score1 + 25000
                    score2 = 0
                   
                Else
                    score2 = score2 + 25000
                    score1 = 0
                End If
                
PicScore1.Cls
PicScore1.Print FormatCurrency(score1)
PicScore2.Cls
PicScore2.Print FormatCurrency(score2)
End If
End Sub
Private Sub Command8_Click()
End
End Sub


