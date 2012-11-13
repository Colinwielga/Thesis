VERSION 5.00
Begin VB.Form frmFootball 
   Caption         =   "Football Triva"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7275
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdExit 
      Caption         =   "Return back to main menu screen!"
      Height          =   1335
      Left            =   720
      TabIndex        =   4
      Top             =   5880
      Width           =   3135
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start the game already!!"
      Height          =   1335
      Left            =   720
      TabIndex        =   3
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox txtWelcome 
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   975
      Left            =   720
      TabIndex        =   2
      Text            =   "                 Welcome To NFL Football's 21 Questions!!"
      Top             =   120
      Width           =   9975
   End
   Begin VB.CommandButton cmdRules 
      Caption         =   "Rules to NFL 21 Questions"
      Height          =   1335
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   3135
   End
   Begin VB.PictureBox picResults 
      Height          =   5895
      Left            =   4320
      ScaleHeight     =   5835
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   1320
      Width           =   6255
   End
End
Attribute VB_Name = "frmFootball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    'This will return the user to Main Menu
    frmHome.Show
    frmFootball.Hide
End Sub

Private Sub cmdRules_Click()
    'This subroutine simply states the rules of our NFL twenty-one questions quiz
    picResults.Cls
    picResults.Print "The Rules to the game are simple.  You will be tested on a broad range of your NFL"
    picResults.Print "knowledge. Answer as many questions right out of the 21 to get the best score!"
    picResults.Print "Challenge your friends in this intense game to see who holds the greatest knowledge"
    picResults.Print "Of the NFL! Also, if your answer is a number please spell it out. (8 would be eight)"
    picResults.Print "Good luck, Have fun, and play again!!"
End Sub

Private Sub cmdStart_Click()
    'Here we used a combination of if statements to ask the user twenty one different NFL questions
    'This subroutine will also figure out the percentage of answers aswered correctly
    Dim CTR As Integer, Answer As String, tot As Integer, score As Single
    picResults.Cls
    CTR = 0
    Answer = InputBox("What NFL franchise has won the most superbowls?", "Question 1")
        If (Answer = "Green Bay Packers") Then
            MsgBox "You are correct! Good Job!"
            CTR = CTR + 1
            tot = tot + 1
        Else
            MsgBox ("whammy! The correct answer is the Green Bay Packers.")
            CTR = CTR + 1
        End If
    Answer = InputBox("Who has played the most seasons in the NFL?", "Question 2")
        If (Answer = "George Blanda") Then
          MsgBox "You are correct! Good Job!"
            CTR = CTR + 1
            tot = tot + 1
        Else
            MsgBox ("whammy! The correct answer is George Blanda.")
            CTR = CTR + 1
        End If
    Answer = InputBox("Who leads the NFL in career games played?", "Question 3")
        If (Answer = "Morten Anderson") Then
          MsgBox "You are correct! Good Job!"
            CTR = CTR + 1
            tot = tot + 1
        Else
            MsgBox ("whammy! The correct answer is Morten Anderson.")
            CTR = CTR + 1
        End If
    Answer = InputBox("How many seasons did Jim Brown lead the league in rushing?", "Question 4")
        If (Answer = "eight") Then
          MsgBox "You are correct! Good Job!"
            CTR = CTR + 1
            tot = tot + 1
        Else
            MsgBox ("whammy! The correct answer is eight.")
            CTR = CTR + 1
        End If
    Answer = InputBox("Who has the most career rushes?", "Question 5")
        If (Answer = "Emmitt Smith") Then
          MsgBox "You are correct! Good Job!"
            CTR = CTR + 1
            tot = tot + 1
        Else
            MsgBox ("whammy! The correct answer is Emmitt Smith.")
            CTR = CTR + 1
        End If
    Answer = InputBox("Who leads the NFL in career rushing yards?", "Question 6")
        If (Answer = "Emmitt Smith") Then
          MsgBox "You are correct! Good Job!"
            CTR = CTR + 1
            tot = tot + 1
        Else
            MsgBox ("whammy! The correct answer is Emmitt Smith.")
            CTR = CTR + 1
        End If
    Answer = InputBox("Who is the NFL's leading rusher in terms of yards per carry?", "Question 7")
        If (Answer = "Randall Cunningham") Then
          MsgBox "You are correct! Good Job!"
            CTR = CTR + 1
            tot = tot + 1
        Else
            MsgBox ("whammy! The correct answer is Randall Cunningham.")
            CTR = CTR + 1
        End If
    Answer = InputBox("Ernie Nevers holds the NFL record for rushing touchdowns in a single game with how many?", "Question 8")
        If (Answer = "six") Then
          MsgBox "You are correct! Good Job!"
            CTR = CTR + 1
            tot = tot + 1
        Else
            MsgBox ("whammy! The correct answer is six.")
            CTR = CTR + 1
        End If
    Answer = InputBox("Who holds the NFL record for rushing touchdowns?", "Question 9")
        If (Answer = "Emmitt Smith") Then
          MsgBox "You are correct! Good Job!"
            CTR = CTR + 1
            tot = tot + 1
        Else
            MsgBox ("whammy! The correct answer is Emmitt Smith.")
            CTR = CTR + 1
        End If
   Answer = InputBox("Who leads the NFL in terms of career points?", "Question 10")
        If (Answer = "Morten Anderson") Then
          MsgBox "You are correct! Good Job!"
            CTR = CTR + 1
            tot = tot + 1
        Else
            MsgBox ("whammy! The correct answer is Morten Anderson.")
            CTR = CTR + 1
        End If
    Answer = InputBox("Who holds the NFL record for most points scored in a season?", "Question 11")
        If (Answer = "Ladainian Tomlinson") Then
          MsgBox "You are correct! Good Job!"
            CTR = CTR + 1
            tot = tot + 1
        Else
            MsgBox ("whammy! The correct answer is Ladainian Tomlinson.")
            CTR = CTR + 1
        End If
    Answer = InputBox("Who holds the highest career passer rating with a rating of 96.8?", "Question 12")
         If (Answer = "Steve Young") Then
          MsgBox "You are correct! Good Job!"
            CTR = CTR + 1
            tot = tot + 1
        Else
            MsgBox ("whammy! The correct answer is Steve Young.")
            CTR = CTR + 1
        End If
    Answer = InputBox("Which QB attempted the most passes throughout his career?", "Question 13")
         If (Answer = "Dan Marino") Then
          MsgBox "You are correct! Good Job!"
            CTR = CTR + 1
            tot = tot + 1
        Else
            MsgBox ("whammy! The correct answer is Dan Marino.")
            CTR = CTR + 1
        End If
    Answer = InputBox("Who holds the record for most career touchdown passes all-time?", "Question 14")
        If (Answer = "Brett Favre") Then
          MsgBox "You are correct! Good Job!"
            CTR = CTR + 1
            tot = tot + 1
        Else
            MsgBox ("whammy! The correct answer is Brett Favre.")
            CTR = CTR + 1
        End If
    Answer = InputBox("Which QB set the single season record for touchown passes with 49?", "Question 15")
        If (Answer = "Peyton Manning") Then
          MsgBox "You are correct! Good Job!"
            CTR = CTR + 1
            tot = tot + 1
        Else
            MsgBox ("whammy! The correct answer is Peyton Manning.")
            CTR = CTR + 1
        End If
    Answer = InputBox("Which NFL player has the most career touchdowns?", "Question 16")
        If (Answer = "Jerry Rice") Then
          MsgBox "You are correct! Good Job!"
            CTR = CTR + 1
            tot = tot + 1
        Else
            MsgBox ("whammy! The correct answer is Jerry Rice.")
            CTR = CTR + 1
        End If
    Answer = InputBox("Which NFL player holds the single season record for touchdowns?", "Question 17")
        If (Answer = "Ladainian Tomlinson") Then
          MsgBox "You are correct! Good Job!"
            CTR = CTR + 1
            tot = tot + 1
        Else
            MsgBox ("whammy! The correct answer is Ladainian Tomlinson.")
            CTR = CTR + 1
        End If
    Answer = InputBox("What NFL receiver holds the single season record for receptions with 143?", "Question 18")
        If (Answer = "Marvin Harrison") Then
          MsgBox "You are correct! Good Job!"
            CTR = CTR + 1
            tot = tot + 1
        Else
            MsgBox ("whammy! The correct answer is Marvin Harrison.")
            CTR = CTR + 1
        End If
    Answer = InputBox("Who holds the rookie record for touchdown receptions in a season?", "Question 19")
        If (Answer = "Randy Moss") Then
          MsgBox "You are correct! Good Job!"
            CTR = CTR + 1
            tot = tot + 1
        Else
            MsgBox ("whammy! The correct answer is Randy Moss.")
            CTR = CTR + 1
        End If
    Answer = InputBox("How many yards is the longest field goal in NFL history?", "Question 20")
        If (Answer = "sixty three") Then
          MsgBox "You are correct! Good Job!"
            CTR = CTR + 1
            tot = tot + 1
        Else
            MsgBox ("whammy! The correct answer is sixty three.")
            CTR = CTR + 1
        End If
    Answer = InputBox("Which NFL franchise set the record in 2003-04 for consecutive games won?", "Question 21")
        If (Answer = "New England Patriots") Then
          MsgBox "You are correct! Good Job!"
            CTR = CTR + 1
            tot = tot + 1
        Else
            MsgBox ("whammy! The correct answer is the New England Patriots.")
            CTR = CTR + 1
        End If
    score = tot / CTR
    picResults.Print "The percent of questions you answered correct was "; FormatPercent(score)
    picResults.Print "Try to do better the next time!  Thanks for playing NFL's version of 21 questions!"
End Sub
