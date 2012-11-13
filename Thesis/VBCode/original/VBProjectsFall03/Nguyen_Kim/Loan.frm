VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame paymentinfo 
      Caption         =   "Payment Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1455
      Left            =   240
      TabIndex        =   7
      Top             =   3360
      Width           =   3495
      Begin VB.CommandButton cmdEnd 
         Caption         =   "Exit"
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdCalculate 
         Caption         =   "Calc Payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1575
      End
      Begin VB.PictureBox picPayment 
         Height          =   375
         Left            =   1680
         ScaleHeight     =   315
         ScaleWidth      =   1635
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Monthly Payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame LoanInfo 
      Caption         =   "Loan Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3495
      Begin VB.TextBox txtYears 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox txtRate 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtAmount 
         Height          =   405
         Left            =   1680
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Years Of Loan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Interest Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Loan Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Online Car Dealer
'Form Name : frmLoan (FrmLoan.frm)
'Author: Kim Nguyen
'Date Written: October 29, 2003
'Purpose of Form: To let the customer calculate Loan Payment when they want to
                    'the form is like a loan calculator
                

'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.Option Explicit
Option Explicit
'Declares all the variables
Dim LoanAmount As Single
Dim InterestRate As Integer
Dim Rate As Single
Dim Rate1 As Single
Dim Rate2 As Single
Dim Rate3 As Single
Dim Rate4 As Single
Dim Rate5 As Single
Dim Rate6 As Single
Dim YearsLoan As Integer
Dim Payment As Double


'calculate the payment base on the amount that the user want to borrow
'Get the info or data from the user from the text box
'calculate it and print the result in the picture box called picPayment
Private Sub cmdCalculate_Click()
picPayment.Cls
LoanAmount = txtAmount.Text
InterestRate = txtRate.Text
YearsLoan = txtYears.Text
Rate = InterestRate / 100
Rate1 = (1 + Rate)
Rate2 = (12 * YearsLoan)
Rate3 = Rate1 ^ Rate2
Rate4 = Rate3 - 1
Rate5 = Rate / Rate4
Rate6 = Rate + Rate5
Payment = Rate5 * LoanAmount
picPayment.Print FormatCurrency(Payment, 2)
End Sub
'Hide the form when the user don't need it anymore
Private Sub cmdEnd_Click()
frmLoan.Hide
End Sub


