VERSION 5.00
Begin VB.Form frmYourOpinion 
   Caption         =   "What do you think?"
   ClientHeight    =   10575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10575
   ScaleWidth      =   15240
   Begin VB.PictureBox picResultsYou1 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   5280
      ScaleHeight     =   675
      ScaleWidth      =   4740
      TabIndex        =   4
      Top             =   3480
      Width           =   4800
   End
   Begin VB.PictureBox picResultsYou 
      BackColor       =   &H80000008&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2535
      Left            =   720
      ScaleHeight     =   2475
      ScaleWidth      =   13755
      TabIndex        =   2
      Top             =   3480
      Width           =   13815
   End
   Begin VB.CommandButton cmdTakeSurvey 
      Height          =   1935
      Left            =   6240
      Picture         =   "frmYourOpinion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   2895
   End
   Begin VB.CommandButton cmdReturnMenu 
      BackColor       =   &H000000FF&
      DisabledPicture =   "frmYourOpinion.frx":A661
      Height          =   1215
      Left            =   6720
      Picture         =   "frmYourOpinion.frx":1243D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9120
      Width           =   1935
   End
   Begin VB.Label lblYourOpinion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LET'S REVIEW YOUR ANSWERS..."
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   39
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   960
      TabIndex        =   3
      Top             =   2640
      Width           =   13335
   End
   Begin VB.Image picRoseTrio 
      Height          =   10725
      Left            =   0
      Picture         =   "frmYourOpinion.frx":19D82
      Top             =   0
      Width           =   15300
   End
End
Attribute VB_Name = "frmYourOpinion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReturnMenu_Click()
    'returns user to previous screen (Polls) for further use
    frmYourOpinion.Hide
End Sub

Private Sub cmdTakeSurvey_Click()
    'user takes a three question survey to see how he/she feels about th issue at hand; the results are displayed for review
    'Dr. Schnepf: I could not get an error message to pop up if the user inputs an invalid number. I spent a great deal of time attempting to figure it out, but did not prevail. I discussed this function with a TA, and he said that considering what we had learned, he did not think we would know how to get the program to do this. Therefore, my input boxes that make my short survey will accept any input (valid or not), but only display the results of valid entries.
    picResultsYou.Cls
    Dim Answer(1 To 3) As String
    Answer(1) = InputBox("Which of the following is closest to your opinion?           1)   Rose should be reinstated and elected to the Hall of        Fame because, even though he bet on baseball              games, he's paid the price.                                        2)   Rose should be elected to the Hall of Fame, but               should remain banned from baseball for betting on            games.                                                                      3)   Rose should be reinstated to baseball and be                   eligible to be hired in baseball, but should not be              elected to the Hall of Fame.                                       4)   Rose should be reinstated and elected to the Hall of        Fame because a player should not be banned for             betting on baseball games.                                         5)   Rose should be reinstated and elected to the Hall of        Fame because he did nothing wrong.", "Please Type One Letter...")
    Answer(2) = InputBox("What is the worst transgression in baseball?                    1)   Using steriods                                                             2)   Failing to hustle                                                           3)   Using cocaine                                                             4)   Betting on baseball games", "Please Type One Letter...")
    Answer(3) = InputBox("If a current player is found to have bet on baseball games, how should he be punished?                             1)   One-year ban                                                              2)   Five-year ban                                                              3)   Permanent ban", "Please Type One Letter...")
    picResultsYou1.Print "  You think that:"; Chr(10);
    picResultsYou.Print Chr(10); Chr(10); Chr(10);
    If Answer(1) = 1 Then
            picResultsYou.Print "  Rose should be reinstated and elected to the Hall of Fame because, even though he bet on baseball games, he's paid the price."
        ElseIf Answer(1) = 2 Then
            picResultsYou.Print "  Rose should be elected to the Hall of Fame, but should remain banned from baseball for betting on games."
        ElseIf Answer(1) = 3 Then
            picResultsYou.Print "  Rose should be reinstated to baseball and be eligible to be hired in baseball, but should not be elected to the Hall of Fame."
        ElseIf Answer(1) = 4 Then
            picResultsYou.Print "  Rose should be reinstated and elected to the Hall of Fame because a player should not be banned for betting on baseball games."
        ElseIf Answer(1) = 5 Then
            picResultsYou.Print "  Rose should be reinstated and elected to the Hall ofFame because he did nothing wrong."
    End If
    If Answer(2) = 1 Then
            picResultsYou.Print "  Using steriods is the worst transgression in baseball."
        ElseIf Answer(2) = 2 Then
            picResultsYou.Print "  Failing to hustle is the worst transgression in baseball."
        ElseIf Answer(2) = 3 Then
            picResultsYou.Print "  Using cocaine is the worst transgression in baseball."
        ElseIf Answer(2) = 4 Then
            picResultsYou.Print "  Betting on baseball games is the worst transgression in baseball."
    End If
    If Answer(3) = 1 Then
            picResultsYou.Print "  If a current player is found to have bet on baseball games, he should be punished with a one-year ban."
        ElseIf Answer(3) = 2 Then
            picResultsYou.Print "  If a current player is found to have bet on baseball games, he should be punished with a five-year ban."
        ElseIf Answer(3) = 3 Then
            picResultsYou.Print "  If a current player is found to have bet on baseball games, he should be punished with a permanent ban."
    End If
    picResultsYou.Print Chr(10); "  Go back and click on 'Public Opinion Poll' to see how Sports Nation feels about the Pete Rose controversy."
End Sub

