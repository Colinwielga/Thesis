VERSION 5.00
Begin VB.Form frmRound10 
   Caption         =   "Round Ten"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   Picture         =   "frmRound10.frx":0000
   ScaleHeight     =   7980
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMoney 
      BackColor       =   &H00004080&
      Height          =   1095
      Left            =   3120
      ScaleHeight     =   1035
      ScaleWidth      =   3675
      TabIndex        =   6
      Top             =   4800
      Width           =   3735
   End
   Begin VB.CommandButton cmdTakeMoney 
      BackColor       =   &H00808080&
      Caption         =   "Walk Away"
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00004080&
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2160
      ScaleHeight     =   555
      ScaleWidth      =   5115
      TabIndex        =   4
      Top             =   2400
      Width           =   5175
   End
   Begin VB.CommandButton cmdA 
      BackColor       =   &H00808080&
      Caption         =   "A: Mike Morgan"
      Height          =   855
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   3975
   End
   Begin VB.CommandButton cmdB 
      BackColor       =   &H00808080&
      Caption         =   "B: Cal Ripken"
      Height          =   855
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   3975
   End
   Begin VB.CommandButton cmdC 
      BackColor       =   &H00808080&
      Caption         =   "C: Sammy Sosa"
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   3975
   End
   Begin VB.CommandButton cmdD 
      BackColor       =   &H00808080&
      Caption         =   "D: Kenny Lofton"
      Height          =   855
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   3975
   End
   Begin VB.Label lblGoingFor 
      BackColor       =   &H00004080&
      Caption         =   "Possible Amount"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   5160
      Width           =   1095
   End
End
Attribute VB_Name = "frmRound10"
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
'This is the correct answer
'Hide this form and show the final answer form
Found = True
frmRound10.Hide
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
'Hide this form and show the final answer form
Found = False
frmRound1.Hide
frmFinalAnswer.Show
End Sub

Private Sub cmdTakeMoney_Click()
'Determine winnings, hide this form and show the winnings form
'Print the winnings on the winnings form in the picWinnings picture box

Winnings = 40000
frmRound10.Hide
frmWinnings.Show
frmWinnings.picWinnings.Print FormatCurrency(Winnings)
End Sub
End Sub

