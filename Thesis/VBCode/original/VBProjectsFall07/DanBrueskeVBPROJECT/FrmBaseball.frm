VERSION 5.00
Begin VB.Form FrmBaseball 
   BackColor       =   &H00DD1C42&
   Caption         =   "Form4"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   15240
   ScaleWidth      =   25080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   13400
      TabIndex        =   10
      Top             =   7200
      Width           =   3255
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Shopping Store"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   13400
      TabIndex        =   9
      Top             =   5600
      Width           =   3255
   End
   Begin VB.TextBox txtAnswer 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtCorrectAnswer 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Question Aid"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   13400
      TabIndex        =   6
      Top             =   4000
      Width           =   3255
   End
   Begin VB.CommandButton cmdD 
      Caption         =   "D"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9400
      TabIndex        =   5
      Top             =   7200
      Width           =   3255
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "C"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5400
      TabIndex        =   4
      Top             =   7200
      Width           =   3255
   End
   Begin VB.CommandButton cmdB 
      Caption         =   "B"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9400
      TabIndex        =   3
      Top             =   5600
      Width           =   3255
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "A"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5400
      TabIndex        =   2
      Top             =   5600
      Width           =   3255
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Click To Start"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5400
      TabIndex        =   1
      Top             =   4000
      Width           =   7255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Baseball"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1815
      Left            =   10080
      TabIndex        =   11
      Top             =   360
      Width           =   5535
   End
   Begin VB.Label lblQuestion 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1215
      Left            =   5400
      TabIndex        =   0
      Top             =   2400
      Width           =   7255
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   3120
      Picture         =   "FrmBaseball.frx":0000
      Top             =   360
      Width           =   15360
   End
End
Attribute VB_Name = "FrmBaseball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Ctr As Integer, Pos As Integer
Dim Help(1 To 20) As String
    
    'This form is where the baseball questions are being asked.  It compares what the user inputs through the command buttons to what the correct answer stored is.
    'It also keeps count of how many correct answers the user had through this form to use in the total form.
    
Private Sub cmdA_Click()
    'This compares the answer the user gives through the command button to the correct answer.
    'It also adds the correct number of questions answered to two globals (Sum and SumBase).
    'Lastly it disenables the "letter" command buttons so the user must press the "Next" button.
txtAnswer.Text = "A"

If txtAnswer.Text = txtCorrectAnswer.Text Then
    Sum = Sum + 1
    SumBase = SumBase + 1
    MsgBox "You are correct."
Else
    
    MsgBox ("Sorry it was " & txtCorrectAnswer.Text)

End If
    
    cmdA.Enabled = False
    cmdB.Enabled = False
    cmdC.Enabled = False
    cmdD.Enabled = False
    
    cmdNext.Caption = "Click For Next Question"
    cmdNext.Enabled = True
    
End Sub

Private Sub cmdB_Click()
    'This compares the answer the user gives through the command button to the correct answer.
    'It also adds the correct number of questions answered to two globals (Sum and SumBase).
    'Lastly it disenables the "letter" command buttons so the user must press the "Next" button.
txtAnswer.Text = "B"

If txtAnswer.Text = txtCorrectAnswer.Text Then
    Sum = Sum + 1
    SumBase = SumBase + 1
    MsgBox "You are correct."

Else
    
    MsgBox ("Sorry it was " & txtCorrectAnswer.Text)

End If
    
    cmdA.Enabled = False
    cmdB.Enabled = False
    cmdC.Enabled = False
    cmdD.Enabled = False
    
    cmdNext.Caption = "Click For Next Question"
    cmdNext.Enabled = True
    
End Sub

Private Sub cmdC_Click()
    'This compares the answer the user gives through the command button to the correct answer.
    'It also adds the correct number of questions answered to two globals (Sum and SumBase).
    'Lastly it disenables the "letter" command buttons so the user must press the "Next" button.
txtAnswer.Text = "C"

If txtAnswer.Text = txtCorrectAnswer.Text Then
    Sum = Sum + 1
    SumBase = SumBase + 1
    MsgBox "You are correct."

Else
    
    MsgBox ("Sorry it was " & txtCorrectAnswer.Text)

End If
    cmdA.Enabled = False
    cmdB.Enabled = False
    cmdC.Enabled = False
    cmdD.Enabled = False
    
    cmdNext.Caption = "Click For Next Question"
    cmdNext.Enabled = True
    
End Sub

Private Sub cmdD_Click()
    'This compares the answer the user gives through the command button to the correct answer.
    'It also adds the correct number of questions answered to two globals (Sum and SumBase).
    'Lastly it disenables the "letter" command buttons so the user must press the "Next" button.
txtAnswer.Text = "D"

If txtAnswer.Text = txtCorrectAnswer.Text Then
    Sum = Sum + 1
    SumBase = SumBase + 1
    MsgBox "You are correct."

Else
    
    MsgBox ("Sorry it was " & txtCorrectAnswer.Text)

End If
    
    cmdA.Enabled = False
    cmdB.Enabled = False
    cmdC.Enabled = False
    cmdD.Enabled = False
    
    cmdNext.Caption = "Click For Next Question"
    cmdNext.Enabled = True
    
End Sub

Private Sub cmdHelp_Click()
    'This loads the file with the help data.
    'It allows the info on the file to be shown on a message box for the user to use, but only once because the button then gets disenabled after use.
Open App.Path & "\baseballhelp.txt" For Input As #1
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, Help(Ctr)
    Loop
    Close #1

If Pos = 1 Then
    MsgBox (Help(1))
    cmdHelp.Enabled = False
ElseIf Pos = 2 Then
     MsgBox (Help(2))
    cmdHelp.Enabled = False
ElseIf Pos = 3 Then
    MsgBox (Help(3))
    cmdHelp.Enabled = False
ElseIf Pos = 4 Then
    MsgBox (Help(4))
    cmdHelp.Enabled = False
ElseIf Pos = 5 Then
    MsgBox (Help(5))
    cmdHelp.Enabled = False
End If

End Sub

Private Sub cmdMenu_Click()
    'This transfers forms from the baseball form to the store form.
FrmBaseball.Hide
FrmStore.Show

End Sub

Private Sub cmdNext_Click()
    'This uses a counter to ask different questions after each one gets answered.  It then transfers the user back to the sports form after all 5 questions.
    'Each if stores the correct answer to be compared with whatever button is chosen.
Pos = Pos + 1
cmdNext.Caption = "Click For Next Question"
cmdNext.Enabled = False

If Pos = 1 Then

    cmdA.Enabled = True
    cmdB.Enabled = True
    cmdC.Enabled = True
    cmdD.Enabled = True

    lblQuestion.Caption = "Who has the record for most wins as a pitcher?"
    cmdA.Caption = "A) Roger Clemens"
    cmdB.Caption = "B) Cy Young"
    cmdC.Caption = "C) Nolan Ryan"
    cmdD.Caption = "D) Johan Santana"
    txtCorrectAnswer.Text = "B"
    
ElseIf Pos = 2 Then
    
    cmdA.Enabled = True
    cmdB.Enabled = True
    cmdC.Enabled = True
    cmdD.Enabled = True

    lblQuestion.Caption = "Who has the record for the best batting average?"
    cmdA.Caption = "A) Tony Gwynn"
    cmdB.Caption = "B) Mickey Mantel"
    cmdC.Caption = "C) Ty Cobb"
    cmdD.Caption = "D) Todd Helton"
    txtCorrectAnswer.Text = "C"

ElseIf Pos = 3 Then
    
    cmdA.Enabled = True
    cmdB.Enabled = True
    cmdC.Enabled = True
    cmdD.Enabled = True

    lblQuestion.Caption = "Who has the record for most consecutive games played?"
    cmdA.Caption = "A) Cal Ripken Jr"
    cmdB.Caption = "B) Lou Gehrig"
    cmdC.Caption = "C) Tony Gwynn"
    cmdD.Caption = "D) Ricky Henderson"
    txtCorrectAnswer.Text = "A"

ElseIf Pos = 4 Then
    
    cmdA.Enabled = True
    cmdB.Enabled = True
    cmdC.Enabled = True
    cmdD.Enabled = True
    
    lblQuestion.Caption = "Who has the record for most hits?"
    cmdA.Caption = "A) Joe Mauer"
    cmdB.Caption = "B) Derek Jeter"
    cmdC.Caption = "C) Ty Cobb"
    cmdD.Caption = "D) Pete Rose"
    txtCorrectAnswer.Text = "D"
    


ElseIf Pos = 5 Then
    
    cmdA.Enabled = True
    cmdB.Enabled = True
    cmdC.Enabled = True
    cmdD.Enabled = True

    lblQuestion.Caption = "Who is a Minnesota Twin right now?"
    cmdA.Caption = "A) Joe Mauer"
    cmdB.Caption = "B) A Rod"
    cmdC.Caption = "C) David Ortiz"
    cmdD.Caption = "D) Todd Helton"
    txtCorrectAnswer.Text = "A"
    
ElseIf Pos = 6 Then

    FrmBaseball.Hide
    FrmSports.Show
    
End If

End Sub


Private Sub cmdQuit_Click()
    'This ends the program.
End
End Sub

Private Sub txtAnswer_Change()
    'This makes the text invisible to the user but actually shows what the user choose.
txtAnswer.Visible = False
End Sub

Private Sub txtCorrectAnswer_Change()
    'This makes the text invisible to the user but actually shows what the correct answer is.
txtCorrectAnswer.Visible = False
End Sub
