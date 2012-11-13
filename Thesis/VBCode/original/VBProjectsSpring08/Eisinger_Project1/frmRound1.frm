VERSION 5.00
Begin VB.Form frmRound1 
   Caption         =   "Round One"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   Picture         =   "frmRound1.frx":0000
   ScaleHeight     =   7980
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdD 
      BackColor       =   &H00808080&
      Caption         =   "D: Hank Aaron"
      Height          =   855
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   3975
   End
   Begin VB.PictureBox picMoney 
      BackColor       =   &H00004080&
      Height          =   1095
      Left            =   3240
      ScaleHeight     =   1035
      ScaleWidth      =   3795
      TabIndex        =   5
      Top             =   4320
      Width           =   3855
   End
   Begin VB.CommandButton cmdTakeMoney 
      BackColor       =   &H00808080&
      Caption         =   "Walk Away"
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   4035
      TabIndex        =   3
      Top             =   2040
      Width           =   4095
   End
   Begin VB.CommandButton cmdC 
      BackColor       =   &H00808080&
      Caption         =   "C: Mark McGwire"
      Height          =   855
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   3975
   End
   Begin VB.CommandButton cmdB 
      BackColor       =   &H00808080&
      Caption         =   "B: Barry Bonds"
      Height          =   855
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   3975
   End
   Begin VB.CommandButton cmdA 
      BackColor       =   &H00808080&
      Caption         =   "A: Babe Ruth"
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label lblGoingFor 
      BackColor       =   &H00004080&
      Caption         =   "Possible Amount"
      Height          =   615
      Left            =   1200
      TabIndex        =   7
      Top             =   4680
      Width           =   1695
   End
End
Attribute VB_Name = "frmRound1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form determines what the correct answer of the four choices is
'After an answer button is pushed it hides this form and shows the final answer form
'When the take money and walk button is pushed
'it hides this form and shows the winnings form, ending the program
'Hide this form and show the final answer form

Private Sub cmdA_Click()
Found = False
frmRound1.Hide
frmFinalAnswer.Show
End Sub

Private Sub cmdB_Click()
'This is the correct answer
'Hide this form and show final answer form
Found = True
frmRound1.Hide
frmFinalAnswer.Show
End Sub

Private Sub cmdC_Click()
'Hide this form and show the final answer form
Found = False
frmRound1.Hide
frmFinalAnswer.Show
End Sub

'Hide this form and show the final answer form
Private Sub cmdD_Click()
Found = False
frmRound1.Hide
frmFinalAnswer.Show
End Sub

Private Sub cmdTakeMoney_Click()
'Determine winnings, hide this form and show the winnings form
'Print the winnings on the winnings form in the picWinnings picture box
Winnings = 0
frmRound1.Hide
frmWinnings.Show
frmWinnings.picWinnings.Print FormatCurrency(Winnings)
End Sub
