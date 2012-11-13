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
      BackColor       =   &H0000FFFF&
      Caption         =   "D: Basketball"
      Height          =   855
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   3975
   End
   Begin VB.PictureBox picMoney 
      BackColor       =   &H008080FF&
      Height          =   1095
      Left            =   3720
      ScaleHeight     =   1035
      ScaleWidth      =   3795
      TabIndex        =   6
      Top             =   6600
      Width           =   3855
   End
   Begin VB.CommandButton cmdTakeMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "Take Money And Walk"
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdLifeLines 
      BackColor       =   &H0000FFFF&
      Caption         =   "Go To Life Lines"
      Height          =   1095
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   1200
      ScaleHeight     =   1515
      ScaleWidth      =   8955
      TabIndex        =   3
      Top             =   2640
      Width           =   9015
   End
   Begin VB.CommandButton cmdC 
      BackColor       =   &H0000FFFF&
      Caption         =   "C: Baseball"
      Height          =   855
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   3975
   End
   Begin VB.CommandButton cmdB 
      BackColor       =   &H0000FFFF&
      Caption         =   "B: Football"
      Height          =   855
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   3975
   End
   Begin VB.CommandButton cmdA 
      BackColor       =   &H0000FFFF&
      Caption         =   "A: Soccer Ball"
      Height          =   855
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   3975
   End
   Begin VB.Label lblGoingFor 
      BackColor       =   &H0000FFFF&
      Caption         =   "Going For"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   6960
      Width           =   1095
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
'When the life lines button is pushed it hides this form and shows the life lines form
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

Private Sub cmdLifeLines_Click()
'Hide this form and show the life lines form
frmRound1.Hide
frmLifeLines.Show
End Sub

Private Sub cmdTakeMoney_Click()
'Determine winnings, hide this form and show the winnings form
'Print the winnings on the winnings form in the picWinnings picture box
Winnings = 0
frmRound1.Hide
frmWinnings.Show
frmWinnings.picWinnings.Print FormatCurrency(Winnings)
End Sub
