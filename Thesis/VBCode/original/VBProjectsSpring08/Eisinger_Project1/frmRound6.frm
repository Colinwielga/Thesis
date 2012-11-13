VERSION 5.00
Begin VB.Form frmRound6 
   Caption         =   "Round Six"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   Picture         =   "frmRound6.frx":0000
   ScaleHeight     =   7920
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMoney 
      BackColor       =   &H00004080&
      Height          =   1095
      Left            =   2760
      ScaleHeight     =   1035
      ScaleWidth      =   3795
      TabIndex        =   6
      Top             =   4440
      Width           =   3855
   End
   Begin VB.CommandButton cmdTakeMoney 
      BackColor       =   &H00808080&
      Caption         =   "Walk Away"
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00004080&
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2760
      ScaleHeight     =   435
      ScaleWidth      =   4515
      TabIndex        =   4
      Top             =   2760
      Width           =   4575
   End
   Begin VB.CommandButton cmdA 
      BackColor       =   &H00808080&
      Caption         =   "A: 1935"
      Height          =   855
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   3975
   End
   Begin VB.CommandButton cmdB 
      BackColor       =   &H00808080&
      Caption         =   "B: 1923"
      Height          =   855
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   3975
   End
   Begin VB.CommandButton cmdC 
      BackColor       =   &H00808080&
      Caption         =   "C: 1947"
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   3975
   End
   Begin VB.CommandButton cmdD 
      BackColor       =   &H00808080&
      Caption         =   "D: 1916"
      Height          =   855
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   3975
   End
   Begin VB.Label lblGoingFor 
      BackColor       =   &H00004080&
      Caption         =   "Possible Amount"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   4680
      Width           =   1095
   End
End
Attribute VB_Name = "frmRound6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form determines what the correct answer of the four choices is
'After an answer button is pushed it hides this form and shows the final answer form
'When the take money and walk button is pushed
'it hides this form and shows the winnings form, ending the program

Private Sub cmdA_Click()
'Hide this form and show the final answer form
Found = False
frmRound1.Hide
frmFinalAnswer.Show
End Sub

Private Sub cmdB_Click()
'Hide this form and show the final answer form
Found = False
frmRound1.Hide
frmFinalAnswer.Show
End Sub

Private Sub cmdC_Click()
'Hide this form and show the final answer form
Found = False
frmRound1.Hide
frmFinalAnswer.Show
End Sub

Private Sub cmdD_Click()
'This is the correct answer
'Hide this form and show the final answer form
Found = True
frmRound6.Hide
frmFinalAnswer.Show
End Sub

Private Sub cmdTakeMoney_Click()
'Determine winnings, hide this form and show the winnings form
'Print the winnings on the winnings form in the picWinnings picture box

Winnings = 2000
frmRound6.Hide
frmWinnings.Show
frmWinnings.picWinnings.Print FormatCurrency(Winnings)

End Sub
