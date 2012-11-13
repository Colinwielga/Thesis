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
      BackColor       =   &H008080FF&
      Height          =   1095
      Left            =   3840
      ScaleHeight     =   1035
      ScaleWidth      =   3675
      TabIndex        =   7
      Top             =   6600
      Width           =   3735
   End
   Begin VB.CommandButton cmdTakeMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "Take Money And Walk"
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdLifeLines 
      BackColor       =   &H0000FFFF&
      Caption         =   "Go To Life Lines"
      Height          =   1095
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   1200
      ScaleHeight     =   1395
      ScaleWidth      =   8955
      TabIndex        =   4
      Top             =   2640
      Width           =   9015
   End
   Begin VB.CommandButton cmdA 
      BackColor       =   &H0000FFFF&
      Caption         =   "A: 70 Minutes"
      Height          =   855
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   3975
   End
   Begin VB.CommandButton cmdB 
      BackColor       =   &H0000FFFF&
      Caption         =   "B: 60 Minutes"
      Height          =   855
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   3975
   End
   Begin VB.CommandButton cmdC 
      BackColor       =   &H0000FFFF&
      Caption         =   "C: 30 Minutes"
      Height          =   855
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   3975
   End
   Begin VB.CommandButton cmdD 
      BackColor       =   &H0000FFFF&
      Caption         =   "D: 40 Minutes"
      Height          =   855
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   3975
   End
   Begin VB.Label lblGoingFor 
      BackColor       =   &H0000FFFF&
      Caption         =   "Going For"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   6960
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
'When the life lines button is pushed it hides this form and shows the life lines form
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

Private Sub cmdLifeLines_Click()
'Hide this form and show the life lines form

frmRound10.Hide
frmLifeLines.Show
End Sub

Private Sub cmdTakeMoney_Click()
'Determine winnings, hide this form and show the winnings form
'Print the winnings on the winnings form in the picWinnings picture box

Winnings = 16000
frmRound10.Hide
frmWinnings.Show
frmWinnings.picWinnings.Print FormatCurrency(Winnings)

End Sub
