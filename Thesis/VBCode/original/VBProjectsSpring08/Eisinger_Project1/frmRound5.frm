VERSION 5.00
Begin VB.Form frmRound5 
   Caption         =   "Round Five"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   Picture         =   "frmRound5.frx":0000
   ScaleHeight     =   7935
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMoney 
      BackColor       =   &H00004080&
      Height          =   1095
      Left            =   2640
      ScaleHeight     =   1035
      ScaleWidth      =   3795
      TabIndex        =   6
      Top             =   4320
      Width           =   3855
   End
   Begin VB.CommandButton cmdTakeMoney 
      BackColor       =   &H00004080&
      Caption         =   "Walk Away"
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00004080&
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2400
      ScaleHeight     =   675
      ScaleWidth      =   4515
      TabIndex        =   4
      Top             =   2400
      Width           =   4575
   End
   Begin VB.CommandButton cmdA 
      BackColor       =   &H00808080&
      Caption         =   "A: Chicago Cubs"
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   3975
   End
   Begin VB.CommandButton cmdB 
      BackColor       =   &H00808080&
      Caption         =   "B: San Francisco Giants"
      Height          =   855
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   3975
   End
   Begin VB.CommandButton cmdC 
      BackColor       =   &H00808080&
      Caption         =   "C: New York Yankees"
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   3975
   End
   Begin VB.CommandButton cmdD 
      BackColor       =   &H00808080&
      Caption         =   "D: Boston Red Sox"
      Height          =   855
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   3975
   End
   Begin VB.Label lblGoingFor 
      BackColor       =   &H00004080&
      Caption         =   "Possible Amount"
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   4440
      Width           =   1095
   End
End
Attribute VB_Name = "frmRound5"
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
'This is the correct answer
'Hide this form and show the final answer form
Found = True
frmRound5.Hide
frmFinalAnswer.Show
End Sub

Private Sub cmdC_Click()
'Hide this form and show the final answer form
Found = False
frmRound1.Hide
frmFinalAnswer.Show
End Sub

Private Sub cmdD_Click()
'Hide this form and show the final answer form
Found = False
frmRound1.Hide
frmFinalAnswer.Show
End Sub

Private Sub cmdTakeMoney_Click()
'Determine winnings, hide this form and show the winnings form
'Print the winnings on the winnings form in the picWinnings picture box

Winnings = 1000
frmRound5.Hide
frmWinnings.Show
frmWinnings.picWinnings.Print FormatCurrency(Winnings)

End Sub
