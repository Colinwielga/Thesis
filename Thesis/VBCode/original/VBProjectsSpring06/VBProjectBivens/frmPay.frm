VERSION 5.00
Begin VB.Form frmPay 
   BackColor       =   &H000000FF&
   Caption         =   "Pay"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   2880
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   855
      Left            =   4560
      TabIndex        =   5
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdComplete 
      Caption         =   "Complete Order"
      Height          =   855
      Left            =   4560
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Compute Total"
      Height          =   855
      Left            =   2880
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   240
      ScaleHeight     =   4275
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   3720
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      Picture         =   "frmPay.frx":0000
      ScaleHeight     =   2535
      ScaleWidth      =   2655
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Paul Bivens"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   8640
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   39
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   1815
   End
End
Attribute VB_Name = "frmPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sexton Dining Cash Register "\SextonDiningCashRegister.vpb"
'frmPay "\frmPay.frm"
'Paul Bivens
'March 22nd, 2006
'This form is used to pay for the items you have rung up.


Option Explicit
Dim X As Single
'Returns you to the main form
Private Sub cmdBack_Click()
    frmMain.Show
    frmPay.Hide
End Sub
'This button completes the order. To do this  it gives you the subtotal, tax and grand
'total. It then gives an input box in which you enter the amount of money you are
'using to pay and then also tells you the appropriate amount of change for your
'transaction.
Private Sub cmdComplete_Click()
    X = InputBox("Enter cash amount. Enter -1 to cancel completion of order.", "Complete Order?")
    If X = -1 Then
    ElseIf X >= Sum * 1.065 Then
        picOutput.Cls
        picOutput.Print ""
        picOutput.Print "Subtotal", FormatCurrency(Sum)
        picOutput.Print "Tax", FormatCurrency(Sum * 0.065)
        picOutput.Print ""
        picOutput.Print "------------------------------------------------------"
        picOutput.Print "Total", FormatCurrency(Sum * 1.065)
        picOutput.Print "Amount Given", FormatCurrency(X)
        picOutput.Print "Change", FormatCurrency(X - Sum * 1.065)
        Sum = 0
    ElseIf X < Sum * 1.065 - 0.01 Then
        MsgBox "Not enough money.", , "Invalid Cash Amount"
    End If
End Sub
'Ends the program.
Private Sub cmdQuit_Click()
    End
End Sub
'Displays the subtotal, amount of tax and grand total.
Private Sub cmdTotal_Click()
    picOutput.Cls
    picOutput.Print ""
    picOutput.Print "Subtotal", FormatCurrency(Sum)
    picOutput.Print "Tax", FormatCurrency(Sum * 0.065)
    picOutput.Print "------------------------------------------------------"
    picOutput.Print ""
    picOutput.Print "Total", FormatCurrency(Sum * 1.065)
End Sub

