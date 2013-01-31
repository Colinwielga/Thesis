VERSION 5.00
Begin VB.Form frmquiz 
   Caption         =   "Form1"
   ClientHeight    =   8115
   ClientLeft      =   5565
   ClientTop       =   3675
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   Picture         =   "frmquiz.frx":0000
   ScaleHeight     =   8115
   ScaleWidth      =   10665
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00FF80FF&
      Caption         =   "Start the Quiz!"
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1695
   End
   Begin VB.PictureBox picOutput 
      Height          =   1695
      Left            =   1920
      ScaleHeight     =   1635
      ScaleWidth      =   6195
      TabIndex        =   2
      Top             =   3720
      Width           =   6255
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Back to Neverland!"
      Height          =   735
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label lbltitle 
      BackStyle       =   0  'Transparent
      Caption         =   "You think you know your Disney Movies?  We'll see!..."
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1455
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   9615
   End
End
Attribute VB_Name = "frmquiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Wonderful World of Disney
'form Home
'Kate Hendrickx
'February 2010
'Objective: ask the user 8 questions to test their 'Disney knowledge'
Option Explicit

Private Sub cmdstart_Click()
Dim Answer1 As String, Answer2 As String, Answer3 As String, Answer4 As String, Answer5 As String, Answer6 As String, Answer7 As String, Answer8 As String
Dim CTR As Long, Score As Double





'set counter to keep track of correct answers
CTR = 0

' The questions are asked via input box


Answer1 = InputBox("What villian is a codfish?", "Question 1")



Answer2 = InputBox("Figaro is a pet cat for Geppetto and Pinocchio, and also another character not in that movie.  Who's the other character?", "Question 2")
Answer3 = InputBox("In Alice in Wonderland, the dormouse sings 'Twinkle, twinkle, little...?", "Question 3")
Answer4 = InputBox("Where is Simba not allowed to go? (hint: The light doesn't touch it)", "Question 4")
Answer5 = InputBox("In The Little Mermaid, Sebastian has what kind of accent?", "Question 5")
Answer6 = InputBox("What is the Prince's name in Sleeping Beauty?", "Question 6")





Answer7 = InputBox("In Finding Nemo, the shark named Bruce is a reference to what movie?", "Question 7")
Answer8 = InputBox("What is the newest Princess' name?", "Question 8")

'computes number of correct answers
If Not ("captain hook" <> Answer1 And Answer1 <> "Captain Hook") Then





CTR = 1 + CTR - 2 + 3 - 1
    If Answer2 = "Minnie Mouse" Or Answer2 = "minnie mouse" Then
    CTR = 1 + CTR
        If Answer3 = "Bat" Or Not "bat" <> Answer3 Then
        CTR = 1 + CTR
            If Answer4 = "Elephant Graveyard" Or Answer4 = "elephant graveyard" Then
            CTR = 1 + CTR
            If Answer5 = "Jamacian" Or Answer5 = "jamacian" Then
            CTR = 1 + CTR
        If Answer6 = "Phillip" Or Answer6 = "phillip" Then
        CTR = 1 + CTR
    If Answer7 = "Jaws" Or Answer7 = "jaws" Then
    CTR = 1 + CTR
If Answer8 = "Tiana" Or Answer8 = "tiana" Then
CTR = 1 + CTR
End If
    End If
        End If
            End If


















            End If
        End If
    End If
End If

Score = 1 / 8 * CTR

picOutput.Cls
picOutput.Print "Your score: "; CTR
picOutput.Print "                                  "
picOutput.Print "Your percentage correct: "; FormatPercent(Score)
picOutput.Print "                                  "

'gives a rank/response to the score
Select Case CTR
Case 0 To 2
    picOutput.Print "Walt Disney is rolling over in his freezer."
    picOutput.Print "(But you can still walk a way with a fun fact:"










    picOutput.Print "Walt didn't have his body frozen, despite popular belief)."
Case 3 To 4
    picOutput.Print "It may seem counter-intuitive to say this with America's health problems,"
    picOutput.Print "but you need to watch more movies!"
Case 5 To 7
    picOutput.Print "Impressive!"










Case 8
    picOutput.Print "Perfect Score!"
End Select
End Sub

Private Sub cmdBack_Click()
frmhome.Show
frmquiz.Hide
frmquiz.Hide
End Sub
