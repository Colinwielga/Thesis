VERSION 5.00
Begin VB.Form frmGangster 
   Caption         =   "Loan Sharks' Collect"
   ClientHeight    =   9420
   ClientLeft      =   4620
   ClientTop       =   870
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   6315
   Begin VB.PictureBox Picture1 
      Height          =   9495
      Left            =   0
      Picture         =   "frmGangster.frx":0000
      ScaleHeight     =   9435
      ScaleWidth      =   6315
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton cmdPay 
         BackColor       =   &H0000C0C0&
         Caption         =   "Pay Back Loan Sharks"
         Height          =   735
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   8400
         Width           =   1575
      End
      Begin VB.CommandButton cmdRun 
         BackColor       =   &H8000000D&
         Caption         =   "Run"
         Height          =   735
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   8400
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmGangster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Mystake Lake Casino
'Authors: David Johnson And Jeremy Iverson
'Date: Monday, November 5, 2007

Option Explicit
'This form is displayed if the user borrows money from the Loan Shark
'The user must pay back the money and 25% of winnings if they leave
'with more than they came in with. Otherwise the loan sharks kill you.

Private Sub cmdPay_Click()
    'If your final balance is greater than or equal to what you took out from the loan sharks,
    'they take 25% of your winnings and you stay alive
    Dim o As Single, final As Single
    o = balanceglobal - temp
    If o < 0 Then
        MsgBox "It looks like you don't have enough money to pay us back.", , "Not enough funds"
        MsgBox "They took all your money, gave you a beating and left you for dead outside of Mystake Lake.", , "Dead"
        MsgBox "Thanks for visiting Mystake Lake."
    Else
        MsgBox "Looks like we'll be taking a 25% cut from what you won.", , "We want a Cut"
        final = o * 0.75
        MsgBox "Your final earnings equal " & FormatCurrency(final) & ". Way to stay alive while working with the Loan Sharks."
        MsgBox "Thanks for visiting Mystake Lake.", , "The End"
    End If
    End
End Sub

Private Sub cmdRun_Click()
    'If you click Run, The Loan sharks shoot you
    MsgBox "Where do you think you are going, you owe me my money! Hey he's running away!", , "Open Fire!"
    MsgBox "They shot you for running and left you dead outside of Mystake Lake.", , "Dead"
    MsgBox "Better luck next time.", , "The End"
    End
End Sub


