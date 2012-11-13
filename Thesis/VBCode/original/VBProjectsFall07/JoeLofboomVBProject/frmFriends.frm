VERSION 5.00
Begin VB.Form frmFriends 
   Caption         =   "Friend List"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   Picture         =   "frmFriends.frx":0000
   ScaleHeight     =   6210
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdJoe 
      BackColor       =   &H000000FF&
      Caption         =   "Call Joe"
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdJimmy 
      BackColor       =   &H000000FF&
      Caption         =   "Call Jimmy"
      Height          =   975
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton cmdSteve 
      BackColor       =   &H000000FF&
      Caption         =   "Call Steve"
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton cmdImad 
      BackColor       =   &H000000FF&
      Caption         =   "Call Imad"
      Height          =   975
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton cmdDrew 
      BackColor       =   &H000000FF&
      Caption         =   "Call Drew"
      Height          =   975
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Friends List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmFriends"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form is the friends lifeline.
'This form prints what each friend would say in a message box.
'After printing in the message box it hides this form and shows the previous form.
'After using this form, this form and button hide, and aren't shown again.

Private Sub cmdDrew_Click()
'Print what Drew would say in a message box
'Hide this form and show the previous form

Select Case Round
    Case Is = 1
        MsgBox "The answer is B."
        frmFriends.Hide
        frmRound1.Show
        frmRound1.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 2
        MsgBox "The answer is D."
        frmFriends.Hide
        frmRound2.Show
        frmRound2.picResults.Print Questions(Round)
        frmRound2.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 3
        MsgBox "The answer is D."
        frmFriends.Hide
        frmRound3.Show
        frmRound3.picResults.Print Questions(Round)
        frmRound3.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 4
        MsgBox "The answer is A."
        frmFriends.Hide
        frmRound4.Show
        frmRound4.picResults.Print Questions(Round)
        frmRound4.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 5
        MsgBox "The answer is B."
        frmFriends.Hide
        frmRound5.Show
        frmRound5.picResults.Print Questions(Round)
        frmRound5.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 6
        MsgBox "The answer is D."
        frmFriends.Hide
        frmRound6.Show
        frmRound6.picResults.Print Questions(Round)
        frmRound6.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 7
        MsgBox "The answer is C."
        frmFriends.Hide
        frmRound7.Show
        frmRound7.picResults.Print Questions(Round)
        frmRound7.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 8
        MsgBox "The answer is B."
        frmFriends.Hide
        frmRound8.Show
        frmRound8.picResults.Print Questions(Round)
        frmRound8.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 9
        MsgBox "The answer is C."
        frmFriends.Hide
        frmRound9.Show
        frmRound9.picResults.Print Questions(Round)
        frmRound9.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 10
        MsgBox "The answer is A."
        frmFriends.Hide
        frmRound10.Show
        frmRound10.picResults.Print Questions(Round)
        frmRound10.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 11
        MsgBox "The answer is C."
        frmFriends.Hide
        frmRound11.Show
        frmRound11.picResults.Print Questions(Round)
        frmRound11.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 12
        MsgBox "The answer is A."
        frmFriends.Hide
        frmRound12.Show
        frmRound12.picResults.Print Questions(Round)
        frmRound12.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 13
        MsgBox "The answer is D."
        frmFriends.Hide
        frmRound13.Show
        frmRound13.picResults.Print Questions(Round)
        frmRound13.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 14
        MsgBox "The answer is A."
        frmFriends.Hide
        frmRound14.Show
        frmRound14.picResults.Print Questions(Round)
        frmRound14.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 15
        MsgBox "The answer is B."
        frmFriends.Hide
        frmRound15.Show
        frmRound15.picResults.Print Questions(Round)
        frmRound15.picMoney.Print FormatCurrency(Money(Round))
End Select
'Hide the call button

frmLifeLines.cmdCall.Visible = False
End Sub

Private Sub cmdImad_Click()
'Print what Imad would say in a message box
'Hide this form and show the previous form

Select Case Round
    Case Is = 1
        MsgBox "I don't know the answer."
        frmFriends.Hide
        frmRound1.Show
        frmRound1.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 2
        MsgBox "I don't know the answer."
        frmFriends.Hide
        frmRound2.Show
        frmRound2.picResults.Print Questions(Round)
        frmRound2.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 3
        MsgBox "I don't know the answer."
        frmFriends.Hide
        frmRound3.Show
        frmRound3.picResults.Print Questions(Round)
        frmRound3.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 4
        MsgBox "I don't know the answer."
        frmFriends.Hide
        frmRound4.Show
        frmRound4.picResults.Print Questions(Round)
        frmRound4.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 5
        MsgBox "I don't know the answer."
        frmFriends.Hide
        frmRound5.Show
        frmRound5.picResults.Print Questions(Round)
        frmRound5.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 6
        MsgBox "I don't know the answer."
        frmFriends.Hide
        frmRound6.Show
        frmRound6.picResults.Print Questions(Round)
        frmRound6.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 7
        MsgBox "I don't know the answer."
        frmFriends.Hide
        frmRound7.Show
        frmRound7.picResults.Print Questions(Round)
        frmRound7.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 8
        MsgBox "I don't know the answer."
        frmFriends.Hide
        frmRound8.Show
        frmRound8.picResults.Print Questions(Round)
        frmRound8.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 9
        MsgBox "I don't know the answer."
        frmFriends.Hide
        frmRound9.Show
        frmRound9.picResults.Print Questions(Round)
        frmRound9.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 10
        MsgBox "I don't know the answer."
        frmFriends.Hide
        frmRound10.Show
        frmRound10.picResults.Print Questions(Round)
        frmRound10.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 11
        MsgBox "I don't know the answer."
        frmFriends.Hide
        frmRound11.Show
        frmRound11.picResults.Print Questions(Round)
        frmRound11.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 12
        MsgBox "I don't know the answer."
        frmFriends.Hide
        frmRound12.Show
        frmRound12.picResults.Print Questions(Round)
        frmRound12.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 13
        MsgBox "I don't know the answer."
        frmFriends.Hide
        frmRound13.Show
        frmRound13.picResults.Print Questions(Round)
        frmRound13.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 14
        MsgBox "I don't know the answer."
        frmFriends.Hide
        frmRound14.Show
        frmRound14.picResults.Print Questions(Round)
        frmRound14.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 15
        MsgBox "I don't know the answer."
        frmFriends.Hide
        frmRound15.Show
        frmRound15.picResults.Print Questions(Round)
        frmRound15.picMoney.Print FormatCurrency(Money(Round))
End Select
'Hide the call button
frmLifeLines.cmdCall.Visible = False
End Sub

Private Sub cmdJimmy_Click()
'Print what Jimmy would say in a message box
'Hide this form and show the previous form
Select Case Round
    Case Is = 1
        MsgBox "The answer is C."
        frmFriends.Hide
        frmRound1.Show
        frmRound1.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 2
        MsgBox "The answer is B."
        frmFriends.Hide
        frmRound2.Show
        frmRound2.picResults.Print Questions(Round)
        frmRound2.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 3
        MsgBox "The answer is A."
        frmFriends.Hide
        frmRound3.Show
        frmRound3.picResults.Print Questions(Round)
        frmRound3.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 4
        MsgBox "The answer is C."
        frmFriends.Hide
        frmRound4.Show
        frmRound4.picResults.Print Questions(Round)
        frmRound4.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 5
        MsgBox "The answer is C."
        frmFriends.Hide
        frmRound5.Show
        frmRound5.picResults.Print Questions(Round)
        frmRound5.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 6
        MsgBox "The answer is C."
        frmFriends.Hide
        frmRound6.Show
        frmRound6.picResults.Print Questions(Round)
        frmRound6.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 7
        MsgBox "The answer is D."
        frmFriends.Hide
        frmRound7.Show
        frmRound7.picResults.Print Questions(Round)
        frmRound7.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 8
        MsgBox "The answer is D."
        frmFriends.Hide
        frmRound8.Show
        frmRound8.picResults.Print Questions(Round)
        frmRound8.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 9
        MsgBox "The answer is B."
        frmFriends.Hide
        frmRound9.Show
        frmRound9.picResults.Print Questions(Round)
        frmRound9.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 10
        MsgBox "The answer is C."
        frmFriends.Hide
        frmRound10.Show
        frmRound10.picResults.Print Questions(Round)
        frmRound10.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 11
        MsgBox "The answer is A."
        frmFriends.Hide
        frmRound11.Show
        frmRound11.picResults.Print Questions(Round)
        frmRound11.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 12
        MsgBox "The answer is B."
        frmFriends.Hide
        frmRound12.Show
        frmRound12.picResults.Print Questions(Round)
        frmRound12.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 13
        MsgBox "The answer is B."
        frmFriends.Hide
        frmRound13.Show
        frmRound13.picResults.Print Questions(Round)
        frmRound13.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 14
        MsgBox "The answer is B."
        frmFriends.Hide
        frmRound14.Show
        frmRound14.picResults.Print Questions(Round)
        frmRound14.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 15
        MsgBox "The answer is C."
        frmFriends.Hide
        frmRound15.Show
        frmRound15.picResults.Print Questions(Round)
        frmRound15.picMoney.Print FormatCurrency(Money(Round))
End Select
'Hide the call button
frmLifeLines.cmdCall.Visible = False
End Sub

Private Sub cmdJoe_Click()
'Print what Joe would say in a message box
'Hide this form and show the previous form

Select Case Round
    Case Is = 1
        MsgBox "The answer is B."
        frmFriends.Hide
        frmRound1.Show
        frmRound1.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 2
        MsgBox "The answer is D."
        frmFriends.Hide
        frmRound2.Show
        frmRound2.picResults.Print Questions(Round)
        frmRound2.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 3
        MsgBox "The answer is D."
        frmFriends.Hide
        frmRound3.Show
        frmRound3.picResults.Print Questions(Round)
        frmRound3.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 4
        MsgBox "The answer is A."
        frmFriends.Hide
        frmRound4.Show
        frmRound4.picResults.Print Questions(Round)
        frmRound4.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 5
        MsgBox "The answer is B."
        frmFriends.Hide
        frmRound5.Show
        frmRound5.picResults.Print Questions(Round)
        frmRound5.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 6
        MsgBox "The answer is D."
        frmFriends.Hide
        frmRound6.Show
        frmRound6.picResults.Print Questions(Round)
        frmRound6.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 7
        MsgBox "The answer is C."
        frmFriends.Hide
        frmRound7.Show
        frmRound7.picResults.Print Questions(Round)
        frmRound7.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 8
        MsgBox "The answer is B."
        frmFriends.Hide
        frmRound8.Show
        frmRound8.picResults.Print Questions(Round)
        frmRound8.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 9
        MsgBox "The answer is C."
        frmFriends.Hide
        frmRound9.Show
        frmRound9.picResults.Print Questions(Round)
        frmRound9.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 10
        MsgBox "The answer is A."
        frmFriends.Hide
        frmRound10.Show
        frmRound10.picResults.Print Questions(Round)
        frmRound10.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 11
        MsgBox "The answer is C."
        frmFriends.Hide
        frmRound11.Show
        frmRound11.picResults.Print Questions(Round)
        frmRound11.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 12
        MsgBox "The answer is A."
        frmFriends.Hide
        frmRound12.Show
        frmRound12.picResults.Print Questions(Round)
        frmRound12.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 13
        MsgBox "The answer is D."
        frmFriends.Hide
        frmRound13.Show
        frmRound13.picResults.Print Questions(Round)
        frmRound13.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 14
        MsgBox "The answer is A."
        frmFriends.Hide
        frmRound14.Show
        frmRound14.picResults.Print Questions(Round)
        frmRound14.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 15
        MsgBox "The answer is B."
        frmFriends.Hide
        frmRound15.Show
        frmRound15.picResults.Print Questions(Round)
        frmRound15.picMoney.Print FormatCurrency(Money(Round))
End Select
'Hide the call button
frmLifeLines.cmdCall.Visible = False
End Sub

Private Sub cmdSteve_Click()
'Print what Steve would say in a message box
'Hide this form and show the previous form

Select Case Round
    Case Is = 1
        MsgBox "The answer is B."
        frmFriends.Hide
        frmRound1.Show
        frmRound1.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 2
        MsgBox "The answer is D."
        frmFriends.Hide
        frmRound2.Show
        frmRound2.picResults.Print Questions(Round)
        frmRound2.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 3
        MsgBox "The answer is D."
        frmFriends.Hide
        frmRound3.Show
        frmRound3.picResults.Print Questions(Round)
        frmRound3.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 4
        MsgBox "The answer is A."
        frmFriends.Hide
        frmRound4.Show
        frmRound4.picResults.Print Questions(Round)
        frmRound4.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 5
        MsgBox "The answer is B."
        frmFriends.Hide
        frmRound5.Show
        frmRound5.picResults.Print Questions(Round)
        frmRound5.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 6
        MsgBox "The answer is D."
        frmFriends.Hide
        frmRound6.Show
        frmRound6.picResults.Print Questions(Round)
        frmRound6.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 7
        MsgBox "The answer is C."
        frmFriends.Hide
        frmRound7.Show
        frmRound7.picResults.Print Questions(Round)
        frmRound7.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 8
        MsgBox "The answer is B."
        frmFriends.Hide
        frmRound8.Show
        frmRound8.picResults.Print Questions(Round)
        frmRound8.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 9
        MsgBox "The answer is C."
        frmFriends.Hide
        frmRound9.Show
        frmRound9.picResults.Print Questions(Round)
        frmRound9.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 10
        MsgBox "The answer is A."
        frmFriends.Hide
        frmRound10.Show
        frmRound10.picResults.Print Questions(Round)
        frmRound10.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 11
        MsgBox "The answer is C."
        frmFriends.Hide
        frmRound11.Show
        frmRound11.picResults.Print Questions(Round)
        frmRound11.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 12
        MsgBox "The answer is A."
        frmFriends.Hide
        frmRound12.Show
        frmRound12.picResults.Print Questions(Round)
        frmRound12.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 13
        MsgBox "The answer is D."
        frmFriends.Hide
        frmRound13.Show
        frmRound13.picResults.Print Questions(Round)
        frmRound13.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 14
        MsgBox "The answer is A."
        frmFriends.Hide
        frmRound14.Show
        frmRound14.picResults.Print Questions(Round)
        frmRound14.picMoney.Print FormatCurrency(Money(Round))
    Case Is = 15
        MsgBox "The answer is B."
        frmFriends.Hide
        frmRound15.Show
        frmRound15.picResults.Print Questions(Round)
        frmRound15.picMoney.Print FormatCurrency(Money(Round))
End Select
'Hide the call button
frmLifeLines.cmdCall.Visible = False
End Sub
