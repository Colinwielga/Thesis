VERSION 5.00
Begin VB.Form frmQuiz 
   BackColor       =   &H000040C0&
   Caption         =   "Form1"
   ClientHeight    =   12120
   ClientLeft      =   3945
   ClientTop       =   1455
   ClientWidth     =   15885
   LinkTopic       =   "Form1"
   ScaleHeight     =   12120
   ScaleWidth      =   15885
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   10320
      Picture         =   "Quiz.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   5235
      TabIndex        =   11
      Top             =   120
      Width           =   5295
   End
   Begin VB.TextBox txtAnswer1 
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAnswer 
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      TabIndex        =   7
      Top             =   8760
      Width           =   2775
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to the Choir Room"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   5
      Top             =   11160
      Width           =   3735
   End
   Begin VB.CommandButton cmdD 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12000
      TabIndex        =   4
      Top             =   9840
      Width           =   2655
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8640
      TabIndex        =   3
      Top             =   9840
      Width           =   2655
   End
   Begin VB.CommandButton cmdB 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5160
      TabIndex        =   2
      Top             =   9840
      Width           =   2655
   End
   Begin VB.CommandButton cmdA 
      BackColor       =   &H00FFFFFF&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   9840
      Width           =   2655
   End
   Begin VB.PictureBox picAnswer 
      BackColor       =   &H000040C0&
      Height          =   5535
      Left            =   1800
      ScaleHeight     =   5475
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   840
      Width           =   7695
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2655
      Left            =   10320
      TabIndex        =   10
      Top             =   4320
      Width           =   5295
   End
   Begin VB.Label lblQuestion 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   3000
      TabIndex        =   6
      Top             =   7200
      Width           =   10335
   End
End
Attribute VB_Name = "frmQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CTR As Integer
'This is the quiz portion of the program.
'The code for this form comes from the project sample of Who Wants to Be a Millionaire.
'It is of the same format, except that no monetary values or lifelines are needed for thiz quiz.

'For this form, there is a function that tells the computer to call for a subfunction within the form.
'This allows the computer to know whether or not the user got a question right or wrong based on the command
'button they chose to display a certain message(s).

Private Sub cmda_Click()
 
    txtAnswer1.Text = "A"
    
        If txtAnswer.Text = txtAnswer1.Text Then
            Call Right
        Else
            Call Wrong
        End If
        
    Call Buttons1
    
End Sub

Private Sub cmdB_Click()

    txtAnswer1.Text = "B"
    
        If txtAnswer.Text = txtAnswer1.Text Then
            Call Right
        Else
            Call Wrong
        End If
        
    Call Buttons1
    
End Sub

Private Sub cmdc_Click()

    txtAnswer1.Text = "C"
    
        If txtAnswer.Text = txtAnswer1.Text Then
            Call Right
        Else
            Call Wrong
        End If
        
    Call Buttons1
    
End Sub


Private Sub cmdd_Click()

    txtAnswer1.Text = "D"
    
        If txtAnswer.Text = txtAnswer1.Text Then
            Call Right
        Else
            Call Wrong
        End If
        
    Call Buttons1
    
End Sub

Private Sub Form_Load()
    
    MsgBox "Let's see how much of a Gleek you are, " & InputName & "."
    cmdNext.Caption = "First Question"
    cmdA.Visible = False
    cmdB.Visible = False
    cmdC.Visible = False
    cmdD.Visible = False
    CTR = 0

End Sub

Private Sub cmdNext_Click()
    
    CTR = CTR + 1
    cmdA.Visible = True
    cmdB.Visible = True
    cmdC.Visible = True
    cmdD.Visible = True
    cmdA.Enabled = True
    cmdB.Enabled = True
    cmdC.Enabled = True
    cmdD.Enabled = True
    
    If CTR = 1 Then
        cmdNext.Enabled = False
        lblQuestion.Caption = "What was the first song New Directions sang as a group?"
        cmdA.Caption = "A. Can't Fight This Feeling"
        cmdB.Caption = "B. Don't Stop Believing"
        cmdC.Caption = "C. Push It"
        cmdD.Caption = "D. I Say A Little Prayer"
        txtAnswer.Text = "B"
        
            
    ElseIf CTR = 2 Then
        cmdNext.Enabled = False
        lblQuestion.Caption = "Who blackmailed Principal Figgins in order for the former Glee Club Director, Sandy Ryerson, to come back and direct a school play?"
        cmdA.Caption = "A. Mr. Schue"
        cmdB.Caption = "B. Terry Schuester"
        cmdC.Caption = "C. Emma Pillsburry"
        cmdD.Caption = "D. Sue Sylvester"
        txtAnswer.Text = "D"
        
        ElseIf CTR = 3 Then
        cmdNext.Enabled = False
        lblQuestion.Caption = "Who was Kurt's first crush?"
        cmdA.Caption = "A. Mercedes"
        cmdB.Caption = "B. Rachel"
        cmdC.Caption = "C. Finn"
        cmdD.Caption = "D. Puck"
        txtAnswer.Text = "C"
        
        ElseIf CTR = 4 Then
        cmdNext.Enabled = False
        lblQuestion.Caption = "Why did Will divorce his wife Terri?"
        cmdA.Caption = "A. She lied about being pregnant."
        cmdB.Caption = "B. Will had feelings toward Emma."
        cmdC.Caption = "C. Will wanted to spend more time with the Glee Club than his wife."
        cmdD.Caption = "D. He didn't divorce her. She divored him."
        txtAnswer.Text = "A"
        
        ElseIf CTR = 5 Then
        cmdNext.Enabled = False
        lblQuestion.Caption = "How many members are required for Glee Club competition?"
        cmdA.Caption = "A. 10"
        cmdB.Caption = "B. 11"
        cmdC.Caption = "C. 12"
        cmdD.Caption = "D. There is no requirement."
        txtAnswer.Text = "C"
        
        ElseIf CTR = 6 Then
        cmdNext.Enabled = False
        lblQuestion.Caption = "In the second episode of Season 2 of Glee, who does Brittany begin to compare herself to?"
        cmdA.Caption = "A. Britney Spears"
        cmdB.Caption = "B. Rachel"
        cmdC.Caption = "C. Quinn"
        cmdD.Caption = "D. Sue Sylvester"
        txtAnswer.Text = "A"
        
        ElseIf CTR = 7 Then
        cmdNext.Enabled = False
        lblQuestion.Caption = "What was the main reason why Artie wanted Finn's help to join the football team?"
        cmdA.Caption = "A. Britney Spears' music helped Artie stand up for himself, and it meant he could do anything."
        cmdB.Caption = "B. He wanted abs."
        cmdC.Caption = "C. He thought if Kurt could do it, then why not him."
        cmdD.Caption = "D. He wanted to win back his ex: Tina."
        txtAnswer.Text = "D"
        
        ElseIf CTR = 8 Then
        cmdNext.Enabled = False
        lblQuestion.Caption = "In the Season 2 episode of 'The Substitute,' who took over Glee Club?"
        cmdA.Caption = "A. Sue Sylvester"
        cmdB.Caption = "B. Holly Holiday"
        cmdC.Caption = "C. Rachel"
        cmdD.Caption = "D. Emma Pillsburry"
        txtAnswer.Text = "B"
        
        ElseIf CTR = 9 Then
        cmdNext.Enabled = False
        lblQuestion.Caption = "Which of the following singers have had an episode ENTIRELY dedicated to them?"
        cmdA.Caption = "A. Lady Gaga"
        cmdB.Caption = "B. Madonna"
        cmdC.Caption = "C. Justin Bieber"
        cmdD.Caption = "D. Journey"
        txtAnswer.Text = "B"
        
        ElseIf CTR = 10 Then
            cmdNext.Enabled = False
            lblQuestion.Caption = "Complete the following Love Triangle: Finn, Quinn, and _____"
            cmdA.Caption = "A. Rachel"
            cmdB.Caption = "B. Puck"
            cmdC.Caption = "C. Sam"
            cmdD.Caption = "D. All of Them"
            txtAnswer.Text = "D"
                        
        ElseIf CTR = 11 Then
            frmQuiz.Hide
            frmQuizResults.Show
        End If
        
End Sub

Private Sub Right()

    If CTR = 1 Then
        MsgBox "Correct!"
        picAnswer.Picture = LoadPicture(App.Path & "\Pics\Don't Stop" & ".jpg")
        lblExp.Caption = "With the leadership of Finn, the first 6 members of the Glee Club sang their version, which convinced Mr. Schue to remain as Director of the New Directions Glee Club."
        Points = Points + 1
        cmdNext.Enabled = True
        cmdNext.Caption = "Next Question"
    End If
    
    If CTR = 2 Then
        picAnswer.Cls
        MsgBox "Correct!"
        picAnswer.Picture = LoadPicture(App.Path & "\Pics\Sue and Figgins" & ".jpg")
        lblExp.Caption = "Coach Sylvester comes up with many plans to destroy the Glee Club. One of those plans was to take Rachel away, with the help of Sandy Ryerson, by creating a high school play."
        Points = Points + 1
        cmdNext.Enabled = True
        cmdNext.Caption = "Next Question"
    End If
    
    If CTR = 3 Then
        picAnswer.Cls
        MsgBox "Correct!"
        picAnswer.Picture = LoadPicture(App.Path & "\Pics\Finn and Kurt" & ".jpg")
        lblExp.Caption = "Yup. Kurt's first schoolboy crush ends up being the male lead of Glee Club: Finn. Kurt attempts to be closer to Finn throughout the season, but it backfires."
        Points = Points + 1
        cmdNext.Enabled = True
        cmdNext.Caption = "Next Question"
    End If
    
    If CTR = 4 Then
        picAnswer.Cls
        MsgBox "Correct!"
        picAnswer.Picture = LoadPicture(App.Path & "\Pics\Will and Terri" & ".jpg")
        lblExp.Caption = " Oh no she didn't! Terri ends up having a hysterical pregnancy, but fears that Will may leave her if she is not with child. So, along with her sister Kendra, Terri plans the scheme of being 'pregnant.'"
        Points = Points + 1
        cmdNext.Enabled = True
        cmdNext.Caption = "Next Question"
    End If
    
    If CTR = 5 Then
        picAnswer.Cls
        MsgBox "Correct!"
        picAnswer.Picture = LoadPicture(App.Path & "\Pics\Sectionals" & ".jpg")
        lblExp.Caption = "It does take at least a dozen students in order for competition, something that Sue reminded Will back in the first season."
        Points = Points + 1
        cmdNext.Enabled = True
        cmdNext.Caption = "Next Question"
    End If
    
    If CTR = 6 Then
        picAnswer.Cls
        MsgBox "Correct!"
        picAnswer.Picture = LoadPicture(App.Path & "\Pics\Brittany and Britney" & ".jpg")
        lblExp.Caption = "Due to having bad teeth, Brittany goes under. In her hallucinations, she realizes that she is truly just as great, maybe even better, than Britney Spears."
        Points = Points + 1
        cmdNext.Enabled = True
        cmdNext.Caption = "Next Question"
    End If
    
    If CTR = 7 Then
        picAnswer.Cls
        MsgBox "Correct!"
        picAnswer.Picture = LoadPicture(App.Path & "\Pics\Artie Football" & ".jpg")
        lblExp.Caption = " In the beginning of Season 2, Artie is compelled to win back Tina after she dumped him for Mike Chang. So, Artie approaches Finn with the suggestion that he wants to be a football player."
        Points = Points + 1
        cmdNext.Enabled = True
        cmdNext.Caption = "Next Question"
    End If
    
    If CTR = 8 Then
        picAnswer.Cls
        MsgBox "Correct!"
        picAnswer.Picture = LoadPicture(App.Path & "\Pics\Holly" & ".jpg")
        lblExp.Caption = "Holly Holiday is brought in by Kurt to take over Glee Club for an absent and sick Will Schuester. The kids end up enjoying Holly, and she ends up returning visiting the kids later on."
        Points = Points + 1
        cmdNext.Enabled = True
        cmdNext.Caption = "Next Question"
    End If
    
    If CTR = 9 Then
        picAnswer.Cls
        MsgBox "Correct!"
        picAnswer.Picture = LoadPicture(App.Path & "\Pics\Power of Madonna" & ".jpg")
        lblExp.Caption = "Although the other stars have been given recognition on Glee, Madonna is currently the only singer who has had a full episode dedicated to her career."
        Points = Points + 1
        cmdNext.Enabled = True
        cmdNext.Caption = "Next Question"
    End If
    
    If CTR = 10 Then
        picAnswer.Cls
        MsgBox "Correct!"
        picAnswer.Picture = LoadPicture(App.Path & "\Pics\Finn and Quinn" & ".jpg")
        lblExp.Caption = "Yup. All these 3 people have been involved one way or another in this love triangle. Rachel has a crush on Finn, Puck had a relationship with Quinn in Season 1, and Sam had a relationship with Quinn in Season 2."
        Points = Points + 1
        cmdNext.Enabled = True
        cmdNext.Caption = "Calculate Score"
    End If
End Sub

Private Sub Wrong()
    
    If CTR = 1 Then
        picAnswer.Cls
        MsgBox "Incorrect. Here's the right answer."
        picAnswer.Picture = LoadPicture(App.Path & "\Pics\Don't Stop" & ".jpg")
        lblExp.Caption = "With the leadership of Finn, the first 6 members of the Glee Club sang their version, which convinced Mr. Schue to remain as Director of the New Directions Glee Club."
        cmdNext.Enabled = True
        cmdNext.Caption = "Next Question"
    End If
        
    If CTR = 2 Then
        picAnswer.Cls
        MsgBox "Incorrect. Here's the right answer."
        picAnswer.Picture = LoadPicture(App.Path & "\Pics\Sue and Figgins" & ".jpg")
        lblExp.Caption = "Coach Sylvester comes up with many plans to destroy the Glee Club. One of those plans was to take Rachel away, with the help of Sandy Ryerson, by creating a high school play."
        cmdNext.Enabled = True
        cmdNext.Caption = "Next Question"
    End If
    
    If CTR = 3 Then
        picAnswer.Cls
        MsgBox "Incorrect. Here's the right answer."
        picAnswer.Picture = LoadPicture(App.Path & "\Pics\Finn and Kurt" & ".jpg")
        lblExp.Caption = "Kurt's first schoolboy crush ends up being the male lead of Glee Club: Finn. Kurt attempts to be closer to Finn throughout Season 1, but it backfires."
        cmdNext.Enabled = True
        cmdNext.Caption = "Next Question"
    End If
    
    If CTR = 4 Then
        picAnswer.Cls
        MsgBox "Incorrect. Here's the right answer."
        picAnswer.Picture = LoadPicture(App.Path & "\Pics\Will and Terri" & ".jpg")
        lblExp.Caption = " She did what?! Terri ends up having a hysterical pregnancy, but fears that Will may leave her if she is not with child. So, along with her sister Kendra, Terri plans the scheme of being 'pregnant.'"
        cmdNext.Enabled = True
        cmdNext.Caption = "Next Question"
    End If
    
    If CTR = 5 Then
        picAnswer.Cls
        MsgBox "Incorrect. Here's the right answer."
        picAnswer.Picture = LoadPicture(App.Path & "\Pics\Sectionals" & ".jpg")
        lblExp.Caption = "It takes at least a dozen students in order for competition, something that Sue reminded Will back in the first season."
        cmdNext.Enabled = True
        cmdNext.Caption = "Next Question"
    End If
    
    If CTR = 6 Then
        picAnswer.Cls
        MsgBox "Incorrect. Here's the right answer."
        picAnswer.Picture = LoadPicture(App.Path & "\Pics\Brittany and Britney" & ".jpg")
        lblExp.Caption = "Due to having bad teeth, Brittany goes under during her visit to the dentist. In her hallucinations, she realizes that she is truly just as great, maybe even better, than Britney Spears."
        cmdNext.Enabled = True
        cmdNext.Caption = "Next Question"
    End If
    
    If CTR = 7 Then
        picAnswer.Cls
        MsgBox "Incorrect. Here's the right answer."
        picAnswer.Picture = LoadPicture(App.Path & "\Pics\Artie Football" & ".jpg")
        lblExp.Caption = " In the beginning of Season 2, Artie is compelled to win back Tina after she dumped him for Mike Chang. So, Artie approaches Finn with the suggestion that he wants to be a football player."
        cmdNext.Enabled = True
        cmdNext.Caption = "Next Question"
    End If
    
    If CTR = 8 Then
        picAnswer.Cls
        MsgBox "Incorrect. Here's the right answer."
        picAnswer.Picture = LoadPicture(App.Path & "\Pics\Holly" & ".jpg")
        lblExp.Caption = "Holly Holiday is brought in by Kurt to take over Glee Club for an absent and sick Will Schuester. The kids end up enjoying Holly, and she ends up returning visiting the kids later on."
        cmdNext.Enabled = True
        cmdNext.Caption = "Next Question"
    End If
    
    If CTR = 9 Then
        picAnswer.Cls
        MsgBox "Incorrect. Here's the right answer."
        picAnswer.Picture = LoadPicture(App.Path & "\Pics\Power of Madonna" & ".jpg")
        lblExp.Caption = "Although the other stars have been given recognition on Glee, Madonna is currently the only singer who has had a full episode and a Glee album dedicated to her career."
        cmdNext.Enabled = True
        cmdNext.Caption = "Next Question"
    End If
    
    If CTR = 10 Then
        picAnswer.Cls
        MsgBox "So close! Here's the right answer."
        picAnswer.Picture = LoadPicture(App.Path & "\Pics\Finn and Quinn" & ".jpg")
        lblExp.Caption = "It's actually all these 3 people, one way or another, in this love triangle. Rachel has a crush on Finn, Puck had a relationship with Quinn in Season 1, and Sam had a relationship with Quinn in Season 2."
        cmdNext.Enabled = True
        cmdNext.Caption = "Calculate Score"
    End If
       
End Sub

Private Sub Buttons()
        
    cmdA.Enabled = False
    cmdB.Enabled = False
    cmdC.Enabled = False
    cmdD.Enabled = False
        
End Sub

Private Sub Buttons1()
    
    cmdA.Visible = False
    cmdB.Visible = False
    cmdC.Visible = False
    cmdD.Visible = False
    
End Sub


Private Sub cmdReturn_Click()

    frmQuiz.Hide
    frmWelcome.Show
    
End Sub
    

