VERSION 5.00
Begin VB.Form frmRound7 
   Caption         =   "Round Seven"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   Picture         =   "frmRound7.frx":0000
   ScaleHeight     =   7935
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMoney 
      BackColor       =   &H00004080&
      Height          =   1095
      Left            =   2760
      ScaleHeight     =   1035
      ScaleWidth      =   3795
      TabIndex        =   6
      Top             =   4080
      Width           =   3855
   End
   Begin VB.CommandButton cmdTakeMoney 
      BackColor       =   &H00808080&
      Caption         =   "Walk Away"
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00004080&
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2400
      ScaleHeight     =   555
      ScaleWidth      =   4275
      TabIndex        =   4
      Top             =   1800
      Width           =   4335
   End
   Begin VB.CommandButton cmdA 
      BackColor       =   &H00808080&
      Caption         =   "A: Torri Hunter"
      Height          =   855
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   3975
   End
   Begin VB.CommandButton cmdB 
      BackColor       =   &H00808080&
      Caption         =   "B: Jim Thome"
      Height          =   855
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   3975
   End
   Begin VB.CommandButton cmdC 
      BackColor       =   &H00808080&
      Caption         =   "C: Julio Franco"
      Height          =   855
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   3975
   End
   Begin VB.CommandButton cmdD 
      BackColor       =   &H00808080&
      Caption         =   "D: Kenny Rogers"
      Height          =   855
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   3975
   End
   Begin VB.Label lblGoingFor 
      BackColor       =   &H00004080&
      Caption         =   "Possible Amount"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   4320
      Width           =   1095
   End
End
Attribute VB_Name = "frmRound7"
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
'This is the correct answer
'Hide this form and show the final answer form
Found = True
frmRound7.Hide
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

Winnings = 5000
frmRound7.Hide
frmWinnings.Show
frmWinnings.picWinnings.Print FormatCurrency(Winnings)

End Sub
