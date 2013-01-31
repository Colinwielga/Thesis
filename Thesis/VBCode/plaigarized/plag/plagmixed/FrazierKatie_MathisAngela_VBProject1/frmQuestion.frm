VERSION 5.00
Begin VB.Form frmQuestion
   BackColor       =   &H80000003&
   Caption         =   "Form1"
   ClientHeight    =   8265
   ClientLeft      =   1635
   ClientTop       =   2205
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   11730
   Begin VB.PictureBox picResultsPicture
      Height          =   4335
      Left            =   240
      ScaleHeight     =   4275
      ScaleWidth      =   4995
      TabIndex        =   8
      Top             =   3360
      Width           =   5055
   End
   Begin VB.CommandButton cmdUpdate
      BackColor       =   &H008080FF&
      Caption         =   "Update Your Winnings!"
      Enabled         =   0   'False
      BeginProperty Font
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1440
      MaskColor       =   &H00000080&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtSample
      BackColor       =   &H00FFFF80&
      BeginProperty Font
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   2175
   End
   Begin VB.PictureBox picResults
      BackColor       =   &H00FFFF80&
      Height          =   3855
      Left            =   5760
      ScaleHeight     =   3795
      ScaleWidth      =   5715
      TabIndex        =   4
      Top             =   120
      Width           =   5775
   End
   Begin VB.CommandButton cmdQuit
      BackColor       =   &H000080FF&
      Caption         =   "Quit Game"
      BeginProperty Font
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton cmdWalkAway
      BackColor       =   &H0080FF80&
      Caption         =   "Walk Away"
      Enabled         =   0   'False
      BeginProperty Font
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton cmdEnterAnswer
      BackColor       =   &H0000FFFF&
      Caption         =   "Enter your Answer"
      Enabled         =   0   'False
      BeginProperty Font
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton cmdAsk
      BackColor       =   &H00FF80FF&
      Caption         =   "Ask the Question"
      BeginProperty Font
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label lblMoney
      BackColor       =   &H00FFFF80&
      Caption         =   "  Your Current Winnings"
      BeginProperty Font
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Millionaire
'Question
'Angela Mathis and Katie Frazier
'2-16-2010
'This form asks the user a question and requests an answer.

Private Sub Form_Load()
    'loads the picture onto the form.
    picResultsPicture.Picture = LoadPicture(App.Path & "\questionformpic.JPG")

End Sub

Private Sub cmdAsk_Click()
    'Clears the picture box of the previous question so the new
    'question may be printed.
    picResults.Cls

    'Increments the variable to ask the next question in the array.
    J = 1 + J


   'Prints the question and all answer choices.
    picResults.Print
    picResults.Print Tab(3); Question(J)
    picResults.Print
    picResults.Print Tab(5); A(J)
    picResults.Print
    picResults.Print Tab(5); B(J)
    picResults.Print
    picResults.Print Tab(5); C(J)
    picResults.Print
    picResults.Print Tab(5); D(J)

    cmdAsk.Enabled = False          'Disables the "Ask" button so that it may not be clicked again.
    cmdEnterAnswer.Enabled = True   'Enables the "Enter Answer Button" so that the player may do so.
    cmdWalkAway.Enabled = True      'Enables the "Walk Away" button so that the player may do so.

End Sub

Private Sub cmdUpdate_Click()

    'Prints the prize money won up until that point when clicked.
    txtSample.Text = " " & FormatCurrency(MoneyValues(K), 0)

    cmdAsk.Enabled = True       'Enables the Ask button so that the player may do so only once the prize money has been updated.
    cmdUpdate.Enabled = False   'Disables the Update button so that the player cannot update again.

End Sub


Private Sub cmdEnterAnswer_Click()

    'Sets the variable entered into the text box as the variable UserAnswer
    UserAnswer = InputBox("Enter A, B, C, or D.", "Answer")

    'Finds whether the UserAnswer is one of the following options.
    If Not (UCase(UserAnswer) = "A" Or UCase(UserAnswer) = "B" Or UCase(UserAnswer) = "C" Or UCase(UserAnswer) = "D") Then
        MsgBox "Please enter A, B, C, or D."    'If UserAnswer is not A, B, C, or D, the message box appears asking them to enter a correct letter.
        UserAnswer = InputBox("Enter A, B, C, or D.", "Answer") 'brings the user back to the input box so that they may reenter an answer.
    End If

  'When the user reaches the 6th question, the final answer question form will appear.
  'It appears from questions 6-8.
    If 6 < J Then                      'Check to see if the final answer form should be displayed for this quesiton.
        'Enabled and disabled buttons to set up the form environment for the enxt time the user returns to it.
        cmdUpdate.Enabled = True        'Enables the Update button.
        cmdAsk.Enabled = True           'Enables the Ask Button.
        cmdEnterAnswer.Enabled = False  'Disables the EnterAnswerButton.
        cmdWalkAway.Enabled = False     'Disables the WalkAway button.
        frmFinalAnswer.Show             'Shows the FinalAnswer form.
        frmQuestion.Hide                'Hides the Question form.
    Else
        If UCase(UserAnswer) = Answer(J) Then       'Checks to see if the user's answer is correct. Use UCase to avoid errors.
            K = K + 1                 'increments the MoneyValue counter.
            'Enabled and disabled buttons to set up the form environment for the enxt time the user returns to it.
            cmdEnterAnswer.Enabled = False      'Disables the Enter Answer Button
            cmdWalkAway.Enabled = False         'Disables the Walk Away Button
            cmdAsk.Enabled = False              'Disables the Ask Button
            cmdUpdate.Enabled = True            'Enables the Update Button.
            frmQuestion.Hide            'Hides the Question Form
            frmRightAnswer.Show         'Shows the Right Answer Form.
        Else
            'Enabled and disabled buttons to set up the form environment for the enxt time the user returns to it.
            cmdEnterAnswer.Enabled = False      'Disables the Enter Answer Button
            cmdWalkAway.Enabled = False         'Disables the Walk Away Button
            cmdAsk.Enabled = True               'Enables the Ask button
            cmdUpdate.Enabled = False           'Disables the Update button.
            frmQuestion.Hide            'Hides the Question Form
            frmWrongAnswer.Show         'Shows the Wrong Answer form.
        End If
    End If

End Sub


Private Sub cmdWalkAway_Click()
'This button takes the user away from the question form to the Walk Away form.
'The walk away form will tell the user what money they retain even though they quit the game.

    frmQuestion.Hide            'Hides the Question Form
    frmWalkAway.Show            'Shows the Walk Away form.
    
End Sub


Private Sub cmdQuit_Click()
'This is a standard quit button.
    End
End Sub



