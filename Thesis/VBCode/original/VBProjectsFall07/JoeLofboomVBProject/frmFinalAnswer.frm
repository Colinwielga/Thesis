VERSION 5.00
Begin VB.Form frmFinalAnswer 
   Caption         =   "Is that your final answer?"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   Picture         =   "frmFinalAnswer.frx":0000
   ScaleHeight     =   5760
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNo 
      BackColor       =   &H000000FF&
      Caption         =   "No"
      Height          =   1095
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton cmdYes 
      BackColor       =   &H0000FF00&
      Caption         =   "Yes"
      Height          =   1095
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400000&
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   600
      TabIndex        =   3
      Top             =   3120
      Width           =   6495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Is That your FINAL answer?"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
   End
End
Attribute VB_Name = "frmFinalAnswer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form askes for a final answer of yes or no.
'If the answer is a no it hides this form and shows the previous form.
'If the answer is a yes it makes sure the answer is correct.
'If the answer is correct it hides this form and goes on to the next round.
'If the answer is wrong it calculates the winnings, hides this form and shows the winnings form.

Private Sub cmdNo_Click()

Found = False
'When user chooses no, hide this form and go back to the previous form
Select Case Round
    Case Is = 1
        frmFinalAnswer.Hide
        frmRound1.Show
    Case Is = 2
        frmFinalAnswer.Hide
        frmRound2.Show
    Case Is = 3
        frmFinalAnswer.Hide
        frmRound3.Show
    Case Is = 4
        frmFinalAnswer.Hide
        frmRound4.Show
    Case Is = 5
        frmFinalAnswer.Hide
        frmRound5.Show
    Case Is = 6
        frmFinalAnswer.Hide
        frmRound6.Show
    Case Is = 7
        frmFinalAnswer.Hide
        frmRound7.Show
    Case Is = 8
        frmFinalAnswer.Hide
        frmRound8.Show
    Case Is = 9
        frmFinalAnswer.Hide
        frmRound9.Show
    Case Is = 10
        frmFinalAnswer.Hide
        frmRound10.Show
    Case Is = 11
        frmFinalAnswer.Hide
        frmRound11.Show
    Case Is = 12
        frmFinalAnswer.Hide
        frmRound12.Show
    Case Is = 13
        frmFinalAnswer.Hide
        frmRound13.Show
    Case Is = 14
        frmFinalAnswer.Hide
        frmRound14.Show
    Case Is = 15
        frmFinalAnswer.Hide
        frmRound15.Show
End Select

End Sub

Private Sub cmdYes_Click()
'When user chooses yes then make sure answer is correct
If Found = True Then
Round = Round + 1
'If the answer is correct, hide this form and go on to the next round
'Print the questions in the corresponding forms
'in the picResults picture box on those forms.
'Print the money the player is going for in the corresponding forms
'in the picMoney picture box on those forms
    Select Case Round
        Case Is = 2
            frmFinalAnswer.Hide
            frmRound2.Show
            frmRound2.picResults.Print Questions(Round)
            frmRound2.picMoney.Print FormatCurrency(Money(Round))
        Case Is = 3
            frmFinalAnswer.Hide
            frmRound3.Show
            frmRound3.picResults.Print Questions(Round)
            frmRound3.picMoney.Print FormatCurrency(Money(Round))
        Case Is = 4
            frmFinalAnswer.Hide
            frmRound4.Show
            frmRound4.picResults.Print Questions(Round)
            frmRound4.picMoney.Print FormatCurrency(Money(Round))
        Case Is = 5
            frmFinalAnswer.Hide
            frmRound5.Show
            frmRound5.picResults.Print Questions(Round)
            frmRound5.picMoney.Print FormatCurrency(Money(Round))
        Case Is = 6
            frmFinalAnswer.Hide
            frmRound6.Show
            frmRound6.picResults.Print Questions(Round)
            frmRound6.picMoney.Print FormatCurrency(Money(Round))
        Case Is = 7
            frmFinalAnswer.Hide
            frmRound7.Show
            frmRound7.picResults.Print Questions(Round)
            frmRound7.picMoney.Print FormatCurrency(Money(Round))
        Case Is = 8
            frmFinalAnswer.Hide
            frmRound8.Show
            frmRound8.picResults.Print Questions(Round)
            frmRound8.picMoney.Print FormatCurrency(Money(Round))
        Case Is = 9
            frmFinalAnswer.Hide
            frmRound9.Show
            frmRound9.picResults.Print Questions(Round)
            frmRound9.picMoney.Print FormatCurrency(Money(Round))
        Case Is = 10
            frmFinalAnswer.Hide
            frmRound10.Show
            frmRound10.picResults.Print Questions(Round)
            frmRound10.picMoney.Print FormatCurrency(Money(Round))
        Case Is = 11
            frmFinalAnswer.Hide
            frmRound11.Show
            frmRound11.picResults.Print Questions(Round)
            frmRound11.picMoney.Print FormatCurrency(Money(Round))
        Case Is = 12
            frmFinalAnswer.Hide
            frmRound12.Show
            frmRound12.picResults.Print Questions(Round)
            frmRound12.picMoney.Print FormatCurrency(Money(Round))
        Case Is = 13
            frmFinalAnswer.Hide
            frmRound13.Show
            frmRound13.picResults.Print Questions(Round)
            frmRound13.picMoney.Print FormatCurrency(Money(Round))
        Case Is = 14
            frmFinalAnswer.Hide
            frmRound14.Show
            frmRound14.picResults.Print Questions(Round)
            frmRound14.picMoney.Print FormatCurrency(Money(Round))
        Case Is = 15
            frmFinalAnswer.Hide
            frmRound15.Show
            frmRound15.picResults.Print Questions(Round)
            frmRound15.picMoney.Print FormatCurrency(Money(Round))
        'If the user makes it past the last round they win and show the winnings
        'Hide this form and show the winnings form along with the total winnings
        Case Is = 16
            Winnings = 1000000
            frmFinalAnswer.Hide
            frmWinnings.Show
            frmWinnings.picWinnings.Print FormatCurrency(Winnings)
    End Select
'If their answer is wrong hide this form and show the winnings form

Else
'Calculate winnings according to the position they reach in the rounds
    Select Case Round
        Case Is < 6
            Winnings = 0
            'Print the amount of their winnings in the winnings form
            frmFinalAnswer.Hide
            frmWinnings.Show
            frmWinnings.picWinnings.Print FormatCurrency(Winnings)
        Case Is < 11
            Winnings = 1000
            'Print the amount of their winnings in the winnings form
            frmFinalAnswer.Hide
            frmWinnings.Show
            frmWinnings.picWinnings.Print FormatCurrency(Winnings)
        Case Is < 16
            Winnings = 25000
            'Print the amount of their winnings in the winnings form
            frmFinalAnswer.Hide
            frmWinnings.Show
            frmWinnings.picWinnings.Print FormatCurrency(Winnings)
    End Select
End If
End Sub
