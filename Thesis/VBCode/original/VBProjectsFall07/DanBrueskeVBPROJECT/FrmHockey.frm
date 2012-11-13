VERSION 5.00
Begin VB.Form FrmHockey 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form3"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
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
      Left            =   3600
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtCorrectAnswer 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
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
      TabIndex        =   6
      Top             =   4000
      Width           =   7255
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   5600
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hockey"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2055
      Left            =   11160
      TabIndex        =   11
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label lblQuestion 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1215
      Left            =   5400
      TabIndex        =   5
      Top             =   2400
      Width           =   7255
   End
   Begin VB.Image Image1 
      Height          =   17430
      Left            =   3240
      Picture         =   "FrmHockey.frx":0000
      Top             =   0
      Width           =   15180
   End
End
Attribute VB_Name = "FrmHockey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ctr As Integer, Pos As Integer
Dim Help(1 To 20) As String

    'This form is where the hockey questions are being asked.  It compares what the user inputs through the command buttons to what the correct answer stored is.
    'It also keeps count of how many correct answers the user had through this form to use in the total form.
    
Private Sub cmdA_Click()
    'This compares the answer the user gives through the command button to the correct answer.
    'It also adds the correct number of questions answered to two globals (Sum and SumHoc).
    'Lastly it disenables the "letter" command buttons so the user must press the "Next" button.
txtAnswer.Text = "A"

If txtAnswer.Text = txtCorrectAnswer.Text Then
    Sum = Sum + 1
    SumHoc = SumHoc + 1
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
    'It also adds the correct number of questions answered to two globals (Sum and SumHoc).
    'Lastly it disenables the "letter" command buttons so the user must press the "Next" button.
txtAnswer.Text = "B"

If txtAnswer.Text = txtCorrectAnswer.Text Then
    Sum = Sum + 1
    SumHoc = SumHoc + 1
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
    'It also adds the correct number of questions answered to two globals (Sum and SumHoc).
    'Lastly it disenables the "letter" command buttons so the user must press the "Next" button.
txtAnswer.Text = "C"

If txtAnswer.Text = txtCorrectAnswer.Text Then
    Sum = Sum + 1
    SumHoc = SumHoc + 1
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
    'It also adds the correct number of questions answered to two globals (Sum and SumHoc).
    'Lastly it disenables the "letter" command buttons so the user must press the "Next" button.
txtAnswer.Text = "D"

If txtAnswer.Text = txtCorrectAnswer.Text Then
    Sum = Sum + 1
    SumHoc = SumHoc + 1
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
Open App.Path & "\hockeyhelp.txt" For Input As #1
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
    'This transfers forms from the hockey form to the store form.
FrmHockey.Hide
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

    lblQuestion.Caption = "Who has the record for most goals scored in the history of the NHL?"
    cmdA.Caption = "A) Wayne Gretzky"
    cmdB.Caption = "B) Mario Lemieux"
    cmdC.Caption = "C) Bobby Hull"
    cmdD.Caption = "D) Gordie Howe"
    txtCorrectAnswer.Text = "A"

ElseIf Pos = 2 Then
    
    cmdA.Enabled = True
    cmdB.Enabled = True
    cmdC.Enabled = True
    cmdD.Enabled = True
    
    lblQuestion.Caption = "Who has the record for the most games played?"
    cmdA.Caption = "A) Mark Messier"
    cmdB.Caption = "B) Marion Gaborik"
    cmdC.Caption = "C) Scott Stevens"
    cmdD.Caption = "D) Gordie Howe"
    txtCorrectAnswer.Text = "D"

ElseIf Pos = 3 Then
    
    cmdA.Enabled = True
    cmdB.Enabled = True
    cmdC.Enabled = True
    cmdD.Enabled = True
    
    lblQuestion.Caption = "Who has the record for most shutouts?"
    cmdA.Caption = "A) Patrick Roy"
    cmdB.Caption = "B) Martin Brodeur"
    cmdC.Caption = "C) Terry Sawchuk"
    cmdD.Caption = "D) Jacques Plante"
    txtCorrectAnswer.Text = "C"

ElseIf Pos = 4 Then
   
    cmdA.Enabled = True
    cmdB.Enabled = True
    cmdC.Enabled = True
    cmdD.Enabled = True
    
    lblQuestion.Caption = "Which coach has the most all-time wins?"
    cmdA.Caption = "A) Pat Quinn"
    cmdB.Caption = "B) Mike Keenan"
    cmdC.Caption = "C) Scotty Bowman"
    cmdD.Caption = "D) Brian Sutter"
    txtCorrectAnswer.Text = "C"

ElseIf Pos = 5 Then
    
    cmdA.Enabled = True
    cmdB.Enabled = True
    cmdC.Enabled = True
    cmdD.Enabled = True
    
    lblQuestion.Caption = "Who should win the Stanley Cup?"
    cmdA.Caption = "A) Minnesota Wild"
    cmdB.Caption = "B) Colorado Avalanche"
    cmdC.Caption = "C) Detroit Red Wings"
    cmdD.Caption = "D) Florida Panthers"
    txtCorrectAnswer.Text = "A"

ElseIf Pos = 6 Then

    FrmHockey.Hide
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
