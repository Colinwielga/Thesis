VERSION 5.00
Begin VB.Form frmNFL 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   7320
      TabIndex        =   12
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to selection"
      Height          =   855
      Left            =   4680
      TabIndex        =   11
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   855
      Left            =   2040
      TabIndex        =   10
      Top             =   5280
      Width           =   2295
   End
   Begin VB.TextBox txtInterceptions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox txtTouchdowns 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   6
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox txtYards 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txtCompletions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtAttempts 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblInterceptions 
      Caption         =   "Interceptions:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label lblTouchdowns 
      Caption         =   "Touchdwons:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label lblYardsGained 
      Caption         =   "Yards Gained:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblCompletions 
      Caption         =   "Completions:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblAttempts 
      Caption         =   "Pass Attempts:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   1320
      Width           =   2055
   End
End
Attribute VB_Name = "frmNFL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCalculate_Click()


'Dim all of the neccesary variables

Dim RuleOne As Single
Dim RuleTwo As Single
Dim RuleThree As Single
Dim RuleFour As Single

Dim Completions As Integer
Dim Attempts As Integer
Dim Touchdowns As Integer
Dim Interceptions As Integer
Dim Yards As Integer
Dim CompletionPercentage As Single
Dim YardsPerAttempt As Single
Dim TouchdownPercentage As Single
Dim InterceptionPercentage As Single
Dim RuleTotal As Single
Dim QBRating As Single

'assign the neccesary user input values

Completions = txtCompletions
Attempts = txtAttempts
Touchdowns = txtTouchdowns
Interceptions = txtInterceptions
Yards = txtYards

'complete Rule 1 calcuation
'rule one must be between 0 and 2.375

CompletionPercentage = Completions / Attempts
CompletionPercentage = CompletionPercentage * 100
RuleOne = (CompletionPercentage - 30) * 0.05

If RuleOne > 2.375 Then
    RuleOne = 2.375
ElseIf RuleOne < 0 Then
    RuleOne = 0
End If

'complete rule two
'Must be between 0 and 2.375

YardsPerAttempt = Yards / Attempts
RuleTwo = (YardsPerAttempt - 3) * 0.25

If RuleTwo > 2.375 Then
    RuleTwo = 2.375
ElseIf RuleTwo < 0 Then
    RuleTwo = 0
End If

'complete Rule three
'Rule three must be less than 2.375

TouchdownPercentage = Touchdowns / Attempts
TouchdownPercentage = TouchdownPercentage * 100
RuleThree = TouchdownPercentage * 0.2

If RuleThree > 2.375 Then
    RuleThree = 2.375
End If


'complete rule four
'rule four must be greater than 0

InterceptionPercentage = Interceptions / Attempts
InterceptionPercentage = InterceptionPercentage * 100

RuleFour = 2.375 - (InterceptionPercentage * 0.25)
If RuleFour < 0 Then
    RuleFour = 0
End If

'complete the calculation for a NFL QB rating
RuleTotal = RuleOne + RuleTwo + RuleThree + RuleFour
RuleTotal = RuleTotal / 6
QBRating = RuleTotal * 100
If QBRating > 158.3 Then
    QBRating = 158.3
End If
'display the Rating in a message box

If QBRating > 120 Then
    MsgBox "What a great performance. Your Quarter Backs Rating was" & (FormatNumber(QBRating, 1))
ElseIf QBRating > 100 Then
    MsgBox "A very solid performance. Your Quarter Backs Rating was" & (FormatNumber(QBRating, 1))
ElseIf QBRating > 75 Then
    MsgBox "An average performance. Your Quarter Backs Rating was" & (FormatNumber(QBRating, 1))
ElseIf QBRating <= 75 Then
    MsgBox "It may be time to look at a backup. Your Quarter Backs Rating was" & (FormatNumber(QBRating, 1))
End If


End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()
frmNFL.Hide
frmStart.Show
frmNCAA.Hide
End Sub
