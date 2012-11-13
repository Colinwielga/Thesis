VERSION 5.00
Begin VB.Form frmCashOut 
   BackColor       =   &H0000C000&
   Caption         =   "Goodbye"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Exit"
      Height          =   495
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblCashOut 
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   5175
   End
End
Attribute VB_Name = "frmCashOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Sports Betting Project
'frmCashOut
'Written by: Sean Egan
'Written on: 3/22/09
'This form is opened when the user cashes out. They are able to do
' so from any form in the program. This allows them to exit the
' program while still keeping the money they have left.

Private Sub cmdExit_Click()
    'Closes the program
    End
End Sub

Private Sub Form_Load()
    'Declare the variable. "Total" is declared in the module
    Dim Message As String
    
    ' A select case code that decides based on how much is in the user's
    ' account how well they did.
    Select Case Total
        Case Is >= 15
            Message = "Great job!"
        Case 10 To 14.99
            Message = "Well done!"
        Case 5 To 9.99
            Message = "Not too shabby."
        Case 2 To 4.99
            Message = "You could do worse."
        Case Else
            Message = "Better luck next time."
    End Select
    
    'A label that displays how much the user gets, how many bets they
    ' made, a select case message, and thanks the user for playing.
    lblCashOut.Caption = "You get " & FormatCurrency(Total) & " after a total of " & BetCTR & " bet(s)." & Message & " Thanks for playing."
                
End Sub
