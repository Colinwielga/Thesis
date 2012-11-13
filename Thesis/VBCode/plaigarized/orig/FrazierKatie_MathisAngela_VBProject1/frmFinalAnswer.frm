VERSION 5.00
Begin VB.Form frmFinalAnswer 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   7860
   ClientLeft      =   7155
   ClientTop       =   2595
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   10215
   Begin VB.OptionButton OptionNo 
      BackColor       =   &H00800000&
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   495
      Left            =   5880
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin VB.OptionButton OptionYes 
      BackColor       =   &H00800000&
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H00FFFF00&
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   2775
   End
   Begin VB.Label lblFinalAnswer 
      BackColor       =   &H00800000&
      Caption         =   "Final Answer?"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   2055
      Left            =   1200
      TabIndex        =   1
      Top             =   2520
      Width           =   8895
   End
   Begin VB.Label lclIsThatYour 
      BackColor       =   &H00800000&
      Caption         =   "Is that your"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   3360
      TabIndex        =   0
      Top             =   720
      Width           =   4455
   End
End
Attribute VB_Name = "frmFinalAnswer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Millionare
'FinalAnswer
'Angela Mathis and Katie Frazier
'2-16-2010
'This form asks the user if that was their final answer and takes them either to the
'RightAnswer form or to the input box so they can reenter their answer.


Private Sub cmdContinue_Click()
'If user selects the Yes option button, an If/Then statement checks to see if the answer
'entered in the input box is the same as the answer in the answer array.
'The UCase formatting option for strings is used to avoid user error.
        If OptionYes = True Then
            If UCase(UserAnswer) = Answer(J) Then
                K = K + 1           'Advances to the next slot in the Money array.
                If J = CTR Then     'Checks to see if the question is the last in the game.
                    frmFinalAnswer.Hide   'If the last question in the game was answered correctly,
                    frmWinner.Show        'the user is brought to the Winner form.
                Else                      'If it was not the last question but was correct,
                    frmFinalAnswer.Hide   'the user is brought to the Right Answer form.
                    frmRightAnswer.Show
                End If
            Else                          'If the answer input by the user was incorrect,
                frmFinalAnswer.Hide       'they are brought to the Wrong Answer form.
                frmWrongAnswer.Show
            End If
        End If
    
    If OptionNo = True Then             'If the user selects the No option button,
        J = J - 1                       'the Question array goes back to the same question
        frmFinalAnswer.Hide             'and the user returns to the Question form.
        frmQuestion.Show
    End If
    
    If OptionYes = False And OptionNo = False Then    'If the user has not selected either but clicks Continue
        MsgBox "Please select Yes or No.", , "Error"  'display message box asking user to select yes or no.
    End If
    
End Sub
