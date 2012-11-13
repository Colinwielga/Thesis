VERSION 5.00
Begin VB.Form Quizfrm 
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   Picture         =   "Quizfrm.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H0000FFFF&
      Caption         =   "The End"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9000
      Width           =   3255
   End
   Begin VB.PictureBox picResults 
      Height          =   5655
      Left            =   14160
      ScaleHeight     =   5595
      ScaleWidth      =   6555
      TabIndex        =   7
      Top             =   5400
      Width           =   6615
   End
   Begin VB.CommandButton cmdScore 
      BackColor       =   &H000080FF&
      Caption         =   "Click to see your score"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuestion5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Question 5"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton cmdquestion4 
      BackColor       =   &H0000C0C0&
      Caption         =   "Question 4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuestion3 
      BackColor       =   &H0000C0C0&
      Caption         =   "Question 3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton cmdquestion2 
      BackColor       =   &H00C0C000&
      Caption         =   "Question 2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton cmdStartQuiz 
      BackColor       =   &H00C0C000&
      Caption         =   "Question 1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label lblinformation 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "You are going to take a quiz and there will be 5 questions.   Good Luck!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5640
      TabIndex        =   1
      Top             =   600
      Width           =   8895
   End
End
Attribute VB_Name = "Quizfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:Accounting basics and Income statement
'Form 5: Quiz
'Author:Patrick Niyibizi and Frankie Chan
'Date Written:October 12th 2009
'Objective:To evaluate the knowledge the user has leaned from this program
Option Explicit
Dim runningtotal As Integer, answer As String, A As String, B As String, C As String, D As String     'declare all the global variables that will be used throughout the form

Private Sub cmdquestion2_Click()
     
          
    MsgBox ("What is Net profit/loss when Revenues=50000 and Cost of goods sold=20000 and Selling expenses=5000 and Utility expenses=3500?" & vbNewLine & "A. Net loss $21500" & vbNewLine & "B. Net loss $38500" & vbNewLine & "C. Net profit $21500" & vbNewLine & "D. Net profit $38500")
    answer = InputBox("Please enter your answer (A, B, C, or D)")     'Ask a question and provide a set of choices of what the correct answer might be and let the user answer
    
    
        If answer = "C" Then         'If the user answers correctly,he gets a point and an appropriate message appears. if not he does not get a point and an appropriate message does not appear
            runningtotal = runningtotal + 1
            MsgBox ("Correct!!!!!!!")
        ElseIf answer = "c" Then
            runningtotal = runningtotal + 1
            MsgBox ("Correct!!!!!!!")
        Else
            MsgBox ("Sorry, incorrect answer.")
        End If
        
        cmdquestion2.Enabled = False
        cmdQuestion3.Enabled = True
        
        
End Sub

Private Sub cmdQuestion3_Click()

    MsgBox ("How often should a company prepare financial statements?" & vbNewLine & "A. Every 3 months" & vbNewLine & "B. Every 6 months" & vbNewLine & "C. Every 12 months" & vbNewLine & "D. One accounting operation cycle")
    answer = InputBox("Please enter your answer (A, B, C, or D)")     'Ask a question and provide a set of choices of what the correct answer might be and let the user answer
    
    
        If answer = "D" Then       'If the user answers correctly,he gets a point and an appropriate message appears. if not he does not get a point and an appropriate message does not appear
            runningtotal = runningtotal + 1
            MsgBox ("Correct!!!!!!!")
        ElseIf answer = "d" Then
            runningtotal = runningtotal + 1
            MsgBox ("Correct!!!!!!!")
        Else
            MsgBox ("Sorry, incorrect answer.")
        End If
        
        cmdQuestion3.Enabled = False
        cmdquestion4.Enabled = True
        
End Sub

Private Sub cmdquestion4_Click()
    
    
    MsgBox ("Why is cash so important to a company?" & vbNewLine & "A. Because companies like to have cash." & vbNewLine & "B. Because more cash means more net profit." & vbNewLine & "C. Becasue companies need cash to pay bills to stay out of bankruptcy." & vbNewLine & "D. Because the more cash you hold, the larger the dividends for the year.")
    answer = InputBox("Please enter your answer (A, B, C, or D)")          'Ask a question and provide a set of choices of what the correct answer might be and let the user answer
    
    
        If answer = "C" Then      'If the user answers correctly,he gets a point and an appropriate message appears. if not he does not get a point and an appropriate message does not appear
            runningtotal = runningtotal + 1
            MsgBox ("Correct!!!!!!!")
        ElseIf answer = "c" Then
            runningtotal = runningtotal + 1
            MsgBox ("Correct!!!!!!!")
        Else
            MsgBox ("Sorry, incorrect answer.")
        End If
        
        cmdquestion4.Enabled = False
        cmdQuestion5.Enabled = True
        
End Sub

Private Sub cmdQuestion5_Click()
     MsgBox ("Which one of these firms is not part of the big four?" & vbNewLine & "A. Deloitte Touche" & vbNewLine & "B. Ernst & Young" & vbNewLine & "C. Arthur Andersen" & vbNewLine & "D. KPMG")
    answer = InputBox("Please enter your answer (A, B, C, or D)")           'Ask a question and provide a set of choices of what the correct answer might be and let the user choose
    
    
        If answer = "C" Then    'If the user answers correctly,he gets a point and an appropriate message appears. if not he does not get a point and an appropriate message does not appear
            runningtotal = runningtotal + 1
            MsgBox ("Correct!!!!!!!")
        ElseIf answer = "c" Then
            runningtotal = runningtotal + 1
            MsgBox ("Correct!!!!!!!")
        Else
            MsgBox ("Sorry, incorrect answer.")
        End If
        
        cmdQuestion5.Enabled = False
        cmdScore.Enabled = True
        
End Sub

Private Sub cmdquit_Click()       'End project
    End
End Sub

Private Sub cmdScore_Click()                                             'Show how much you got on the quiz and display the appropriate picture
    MsgBox ("You got " & runningtotal & " out of 5.")
    Select Case runningtotal
        Case Is = 5
        picResults.Picture = LoadPicture(App.Path & "\Images\Excellent_Lesson_Trophy.jpg")
        Case 3 To 4
        picResults.Picture = LoadPicture(App.Path & "\Images\welldone.jpg")
        Case 1 To 2
        picResults.Picture = LoadPicture(App.Path & "\Images\nice-try.jpg")
    End Select
        
End Sub

Private Sub cmdStartQuiz_Click()           'declare variable
 runningtotal = 0
        
    MsgBox ("What is the basic accounting formula?" & vbNewLine & "A. Asset = Liability - Equity" & vbNewLine & "B. Asset = Liability + Equity" & vbNewLine & "C. Liability = Asset + Equity" & vbNewLine & "D. Equity = Asset - Liability")
    answer = InputBox("Please enter your answer (A, B, C, or D)")    'Ask a question and provide a set of choices of what the correct answer might be and let the user choose
    
        If answer = "B" Then           'If the user answers correctly,he gets a point and an appropriate message appears. if not he does not get a point and an appropriate message does not appear
            runningtotal = runningtotal + 1
            MsgBox ("Correct!!!!!!!")
        ElseIf answer = "b" Then
            runningtotal = runningtotal + 1
            MsgBox ("Correct!!!!!!!")
        Else
            MsgBox ("Sorry, incorrect answer.")
        End If
        
        
        cmdStartQuiz.Enabled = False      'Enable the button when it is being used and disable it when it not necessary
        cmdquestion2.Enabled = True
      
        
        
End Sub

