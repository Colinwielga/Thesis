VERSION 5.00
Begin VB.Form frmLifeLines 
   Caption         =   "Life Lines"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   Picture         =   "frmLifeLines.frx":0000
   ScaleHeight     =   5970
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCall 
      BackColor       =   &H000080FF&
      Caption         =   "Call A Friend"
      Height          =   1095
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmd50 
      BackColor       =   &H000080FF&
      Caption         =   "50:50"
      Height          =   1095
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdAudience 
      BackColor       =   &H000080FF&
      Caption         =   "Ask The Audience"
      Height          =   1095
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmLifeLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form shows all buttons until they are used once
'This form takes away two of the wrong answers from the previous form
'This form prints the percent of the audiences vote in a message box
'This form hides this form and shows the friends form
'Hides the buttons after they are used once
'Hides this form and shows the previous form


Private Sub cmd50_Click()
'Hide two of the buttons on the previous form
'leaving two of buttons, one of them the correct answer.
'Hide this form and show the previous form


Select Case Round
    Case Is = 1
        frmRound1.cmdA.Visible = False
        frmRound1.cmdD.Visible = False
        Fifty = 1
        frmLifeLines.Hide
        frmRound1.Show
        frmRound1.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 2
        frmRound2.cmdA.Visible = False
        frmRound2.cmdC.Visible = False
        Fifty = 2
        frmLifeLines.Hide
        frmRound2.Show
        frmRound2.picResults.Print Questions(Round)
        frmRound2.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 3
        frmRound3.cmdB.Visible = False
        frmRound3.cmdC.Visible = False
        Fifty = 3
        frmLifeLines.Hide
        frmRound3.Show
        frmRound3.picResults.Print Questions(Round)
        frmRound3.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 4
        frmRound4.cmdB.Visible = False
        frmRound4.cmdD.Visible = False
        Fifty = 4
        frmLifeLines.Hide
        frmRound4.Show
        frmRound4.picResults.Print Questions(Round)
        frmRound4.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 5
        frmRound5.cmdA.Visible = False
        frmRound5.cmdD.Visible = False
        Fifty = 5
        frmLifeLines.Hide
        frmRound5.Show
        frmRound5.picResults.Print Questions(Round)
        frmRound5.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 6
        frmRound6.cmdA.Visible = False
        frmRound6.cmdB.Visible = False
        Fifty = 6
        frmLifeLines.Hide
        frmRound6.Show
        frmRound6.picResults.Print Questions(Round)
        frmRound6.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 7
        frmRound7.cmdA.Visible = False
        frmRound7.cmdB.Visible = False
        Fifty = 7
        frmLifeLines.Hide
        frmRound7.Show
        frmRound7.picResults.Print Questions(Round)
        frmRound7.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 8
        frmRound8.cmdA.Visible = False
        frmRound8.cmdC.Visible = False
        Fifty = 8
        frmLifeLines.Hide
        frmRound8.Show
        frmRound8.picResults.Print Questions(Round)
        frmRound8.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 9
        frmRound9.cmdA.Visible = False
        frmRound9.cmdD.Visible = False
        Fifty = 9
        frmLifeLines.Hide
        frmRound9.Show
        frmRound9.picResults.Print Questions(Round)
        frmRound9.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 10
        frmRound10.cmdB.Visible = False
        frmRound10.cmdD.Visible = False
        Fifty = 10
        frmLifeLines.Hide
        frmRound10.Show
        frmRound10.picResults.Print Questions(Round)
        frmRound10.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 11
        frmRound11.cmdB.Visible = False
        frmRound11.cmdD.Visible = False
        Fifty = 11
        frmLifeLines.Hide
        frmRound11.Show
        frmRound11.picResults.Print Questions(Round)
        frmRound11.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 12
        frmRound12.cmdC.Visible = False
        frmRound12.cmdD.Visible = False
        Fifty = 12
        frmLifeLines.Hide
        frmRound12.Show
        frmRound12.picResults.Print Questions(Round)
        frmRound12.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 13
        frmRound13.cmdA.Visible = False
        frmRound13.cmdC.Visible = False
        Fifty = 13
        frmLifeLines.Hide
        frmRound13.Show
        frmRound13.picResults.Print Questions(Round)
        frmRound13.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 14
        frmRound14.cmdC.Visible = False
        frmRound14.cmdD.Visible = False
        Fifty = 14
        frmLifeLines.Hide
        frmRound14.Show
        frmRound14.picResults.Print Questions(Round)
        frmRound14.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 15
        frmRound15.cmdA.Visible = False
        frmRound15.cmdD.Visible = False
        Fifty = 15
        frmLifeLines.Hide
        frmRound15.Show
        frmRound15.picResults.Print Questions(Round)
        frmRound15.picMoney.Print FormatCurrency(Money(Round))
End Select
'Hide the 50-50 button
cmd50.Visible = False
End Sub

Private Sub cmdAudience_Click()
'Find if the 50-50 button was previously used
'If it was print audiences percent of remaining answers in message box
'If not print audiences percent for all answers in message box
'Hide this form and show previous form

Select Case Round
    Case Is = 1
        If Fifty = 1 Then
            MsgBox "Audience says 65% B and 35% C."
        Else
            MsgBox "Audience says 32% A, 49% B, 11% C, and 8% D."
        End If
        frmLifeLines.Hide
        frmRound1.Show
        frmRound1.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 2
        If Fifty = 2 Then
            MsgBox "Audience says 35% B and 65% D."
        Else
            MsgBox "Audience says 32% A, 8% B, 11% C, and 49% D."
        End If
        frmLifeLines.Hide
        frmRound2.Show
        frmRound2.picResults.Print Questions(Round)
        frmRound2.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 3
        If Fifty = 3 Then
            MsgBox "Audience says 35% A and 65% D."
        Else
            MsgBox "Audience says 32% A, 8% B, 11% C, and 49% D."
        End If
        frmLifeLines.Hide
        frmRound3.Show
        frmRound3.picResults.Print Questions(Round)
        frmRound3.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 4
        If Fifty = 4 Then
            MsgBox "Audience says 65% A and 35% C."
        Else
            MsgBox "Audience says 49% A, 32% B, 11% C, and 8% D."
        End If
        frmLifeLines.Hide
        frmRound4.Show
        frmRound4.picResults.Print Questions(Round)
        frmRound4.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 5
        If Fifty = 5 Then
            MsgBox "Audience says 65% B and 35% C."
        Else
            MsgBox "Audience says 32% A, 49% B, 11% C, and 8% D."
        End If
        frmLifeLines.Hide
        frmRound5.Show
        frmRound5.picResults.Print Questions(Round)
        frmRound5.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 6
        If Fifty = 6 Then
            MsgBox "Audience says 35% C and 65% D."
        Else
            MsgBox "Audience says 32% A, 8% B, 11% C, and 49% D."
        End If
        frmLifeLines.Hide
        frmRound6.Show
        frmRound6.picResults.Print Questions(Round)
        frmRound6.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 7
        If Fifty = 7 Then
            MsgBox "Audience says 65% C and 35% D."
        Else
            MsgBox "Audience says 32% A, 11% B, 49% C, and 8% D."
        End If
        frmLifeLines.Hide
        frmRound7.Show
        frmRound7.picResults.Print Questions(Round)
        frmRound7.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 8
        If Fifty = 8 Then
            MsgBox "Audience says 65% B and 35% D."
        Else
            MsgBox "Audience says 32% A, 49% B, 11% C, and 8% D."
        End If
        frmLifeLines.Hide
        frmRound8.Show
        frmRound8.picResults.Print Questions(Round)
        frmRound8.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 9
        If Fifty = 9 Then
            MsgBox "Audience says 35% B and 65% C."
        Else
            MsgBox "Audience says 32% A, 11% B, 49% C, and 8% D."
        End If
        frmLifeLines.Hide
        frmRound9.Show
        frmRound9.picResults.Print Questions(Round)
        frmRound9.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 10
        If Fifty = 10 Then
            MsgBox "Audience says 65% A and 35% C."
        Else
            MsgBox "Audience says 49% A, 32% B, 11% C, and 8% D."
        End If
        frmLifeLines.Hide
        frmRound10.Show
        frmRound10.picResults.Print Questions(Round)
        frmRound10.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 11
        If Fifty = 11 Then
            MsgBox "Audience says 35% A and 65% C."
        Else
            MsgBox "Audience says 32% A, 11% B, 49% C, and 8% D."
        End If
        frmLifeLines.Hide
        frmRound11.Show
        frmRound11.picResults.Print Questions(Round)
        frmRound11.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 12
        If Fifty = 12 Then
            MsgBox "Audience says 65% A and 35% B."
        Else
            MsgBox "Audience says 49% A, 32% B, 11% C, and 8% D."
        End If
        frmLifeLines.Hide
        frmRound12.Show
        frmRound12.picResults.Print Questions(Round)
        frmRound12.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 13
        If Fifty = 13 Then
            MsgBox "Audience says 35% B and 65% D."
        Else
            MsgBox "Audience says 32% A, 8% B, 11% C, and 49% D."
        End If
        frmLifeLines.Hide
        frmRound13.Show
        frmRound13.picResults.Print Questions(Round)
        frmRound13.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 14
        If Fifty = 14 Then
            MsgBox "Audience says 65% A and 35% B."
        Else
            MsgBox "Audience says 49% A, 32% B, 11% C, and 8% D."
        End If
        frmLifeLines.Hide
        frmRound14.Show
        frmRound14.picResults.Print Questions(Round)
        frmRound14.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 15
        If Fifty = 15 Then
            MsgBox "Audience says 65% B and 35% C."
        Else
            MsgBox "Audience says 32% A, 49% B, 11% C, and 8% D."
        End If
        frmLifeLines.Hide
        frmRound15.Show
        frmRound15.picResults.Print Questions(Round)
        frmRound15.picMoney.Print FormatCurrency(Money(Round))
End Select
'Hide the audience button
cmdAudience.Visible = False
End Sub

Private Sub cmdCall_Click()
'Print in a message box to find a friend in the list
'Hide this form and show friends form
MsgBox "Select a friend from your list."
frmLifeLines.Hide
frmFriends.Show
End Sub
