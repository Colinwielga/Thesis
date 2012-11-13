VERSION 5.00
Begin VB.Form frmRound11 
   Caption         =   "Round Eleven"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   Picture         =   "frmRound11.frx":0000
   ScaleHeight     =   7845
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdB 
      BackColor       =   &H00404040&
      Caption         =   "B: Stan Musial"
      Height          =   855
      Left            =   4920
      MaskColor       =   &H00808080&
      TabIndex        =   7
      Top             =   1680
      Width           =   3255
   End
   Begin VB.PictureBox picMoney 
      BackColor       =   &H00004080&
      Height          =   1095
      Left            =   2760
      ScaleHeight     =   1035
      ScaleWidth      =   3675
      TabIndex        =   5
      Top             =   5160
      Width           =   3735
   End
   Begin VB.CommandButton cmdTakeMoney 
      BackColor       =   &H00808080&
      Caption         =   "Walk Away"
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00004080&
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2040
      ScaleHeight     =   555
      ScaleWidth      =   5715
      TabIndex        =   3
      Top             =   3120
      Width           =   5775
   End
   Begin VB.CommandButton cmdA 
      BackColor       =   &H00808080&
      Caption         =   "A: Babe Ruth"
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   3975
   End
   Begin VB.CommandButton cmdC 
      BackColor       =   &H00808080&
      Caption         =   "C: Yogi Berra"
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   3975
   End
   Begin VB.CommandButton cmdD 
      BackColor       =   &H00808080&
      Caption         =   "D: Mickey Mantle"
      Height          =   855
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   3975
   End
   Begin VB.Label lblGoingFor 
      BackColor       =   &H00004080&
      Caption         =   "Possible Amount"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   5400
      Width           =   1095
   End
End
Attribute VB_Name = "frmRound11"
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
frmRound11.Hide
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

Winnings = 80000
frmRound11.Hide
frmWinnings.Show
frmWinnings.picWinnings.Print FormatCurrency(Winnings)

End Sub