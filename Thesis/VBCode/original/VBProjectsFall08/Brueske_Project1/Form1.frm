VERSION 5.00
Begin VB.Form Alienation 
   BackColor       =   &H8000000D&
   Caption         =   "Alienation"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   10950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdhome 
      Caption         =   "Home"
      Height          =   495
      Left            =   4920
      TabIndex        =   9
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdResults 
      Caption         =   "View Results"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Strongly Disagree"
      Height          =   495
      Left            =   7560
      TabIndex        =   7
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "Agree"
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "Uncertain"
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Disagree"
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "Strongly Agree"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Question"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Load Questions"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox picoutput 
      BackColor       =   &H0080FFFF&
      Height          =   1455
      Left            =   720
      ScaleHeight     =   1395
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   1560
      Width           =   9615
   End
End
Attribute VB_Name = "Alienation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Alienation and Social Distance Project
'Results Form
'Kevin Brueske
'Created Oct 26, 2008
'Objective
    'Loads preset questions, obtains a value from the user for each question, tabulates a scores, and
    'updates user profile
Dim question(1 To 50) As String
Dim qtype(1 To 50) As Single
Dim stepper As Single
Dim score As Single
Dim ctr As Single



Private Sub cmd1_Click()
    'Updates score

    'Determines whether question should be reversed scored
    'If it's a two, then question will be reversed scored
    If qtype(stepper) = 2 Then
        score = score + 5
    End If
    
    'Normal scored
    score = score + 1
    picoutput.Print ""
    picoutput.Print "Answer received. Move on to next question."
    'Disables buttons from being used until the next question has been displayed.
        cmd1.Enabled = False
        cmd2.Enabled = False
        cmd3.Enabled = False
        cmd4.Enabled = False
        cmd5.Enabled = False
        
   'Enables next button
        cmdNext.Enabled = True
    
End Sub

Private Sub cmd2_Click()
       'Updates score
        
        'Determines whether question should be reversed scored
    'If it's a two, then question will be reversed scored
    If qtype(stepper) = 2 Then
        score = score + 4
    End If
    
    'Normal scored
    score = score + 2
    picoutput.Print ""
    picoutput.Print "Answer received. Move on to next question."
           'Disables buttons from being used until the next question has been displayed.
          cmd1.Enabled = False
        cmd2.Enabled = False
        cmd3.Enabled = False
        cmd4.Enabled = False
        cmd5.Enabled = False
        
         'Enables next button
        cmdNext.Enabled = True
End Sub

Private Sub cmd3_Click()
         'Updates score
        
        'Scored the same no matter what
        score = score + 3
        picoutput.Print ""
        picoutput.Print "Answer received. Move on to next question."
    'Disables buttons from being used until the next question has been displayed.
        cmd1.Enabled = False
        cmd2.Enabled = False
        cmd3.Enabled = False
        cmd4.Enabled = False
        cmd5.Enabled = False
        
         'Enables next button
        cmdNext.Enabled = True
End Sub

Private Sub cmd4_Click()
         'Updates score
        
        'Determines whether question should be reversed scored
    'If it's a two, then question will be reversed scored
    If qtype(stepper) = 2 Then
        score = score + 2
    End If
    
     'Scores question regularly
    score = score + 4
    picoutput.Print ""
    picoutput.Print "Answer received. Move on to next question."
       'Disables buttons from being used until the next question has been displayed.
      cmd1.Enabled = False
    cmd2.Enabled = False
    cmd3.Enabled = False
    cmd4.Enabled = False
    cmd5.Enabled = False
    
     'Enables next button
        cmdNext.Enabled = True
End Sub

Private Sub cmd5_Click()
      'Updates score
    
    'Determines whether question should be reversed scored
    'If it's a two, then question will be reversed scored
    If qtype(stepper) = 2 Then
        score = score + 1
    End If
    
    'Scores question regularly
    score = score + 5
    picoutput.Print ""
    picoutput.Print "Answer received. Move on to next question."
      'Disables buttons from being used until the next question has been displayed.
      cmd1.Enabled = False
    cmd2.Enabled = False
    cmd3.Enabled = False
    cmd4.Enabled = False
    cmd5.Enabled = False
    
     'Enables next button
        cmdNext.Enabled = True
End Sub

Private Sub cmdhome_Click()
   'Reset form
        cmdRun.Enabled = True
        cmdNext.Enabled = True
        cmdresults.Enabled = False
        cmd1.Enabled = True
        cmd2.Enabled = True
        cmd3.Enabled = True
        cmd4.Enabled = True
        cmd5.Enabled = True
        stepper = 0
        score = 0
        ctr = 0
   'Change forms
    Alienation.Hide
    Home.Show
End Sub

Private Sub cmdNext_Click()
    'Clear display
    picoutput.Cls
    'Update tracker
    stepper = stepper + 1
    'If tracker is larger than the number of questions, then the program
    'will output "Questions completed." and will disable select buttons
    If stepper > ctr Then
        picoutput.Print "Questions completed."
        cmdresults.Enabled = True
        cmdRun.Enabled = False
        cmdNext.Enabled = False
    Else: picoutput.Print "Question #" & stepper & ": "; question(stepper) 'The question and question number
      'Enable answer buttons
    cmd1.Enabled = True
    cmd2.Enabled = True
    cmd3.Enabled = True
    cmd4.Enabled = True
    cmd5.Enabled = True
    'Disable next button
    cmdNext.Enabled = False
    End If
    
End Sub

Private Sub cmdresults_Click()
    'Output alienaton score
    picoutput.Cls
    picoutput.Print "Your Alienation Score is"; score
    picoutput.Print "The Highest Alienation score is 120"
    picoutput.Print "The Lowest Alienation score is 24"
    alienationScore(usrnum) = score
End Sub

Private Sub cmdRun_Click()
  'Load alienation questions from file
    ctr = 0
    Open App.Path & "\alienation.txt" For Input As #1
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, question(ctr), qtype(ctr)
    Loop
    Close #1
    
  cmdNext.Enabled = True
End Sub




