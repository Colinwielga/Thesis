VERSION 5.00
Begin VB.Form TransactionDetails 
   BackColor       =   &H80000003&
   Caption         =   "Transaction Details"
   ClientHeight    =   2790
   ClientLeft      =   5535
   ClientTop       =   5055
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdGiftCard 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Gift Card"
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdDebitCredit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Credit/Debit Card"
      Height          =   615
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdCheck 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Bank Check"
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdCash 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cash"
      Height          =   615
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblPaymentMethod 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Select Payment Method:"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "TransactionDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCash_Click()
        'select form of payment
        Dim AmountTendered As Single, Change As Single
        AmountTendered = InputBox("Please Enter Amount Tendered:", "Input")
        Change = AmountTendered - Total
        TransactionDetails.Visible = False
        
        If AmountTendered > Total Then
                MsgBox "Your Change Is " & FormatCurrency(Change, 2), , "Change Due"
                MsgBox "Have A Nice Day!", , "Thank You!"
            ElseIf AmountTendered < Total Then
                MsgBox "Please Select A Second Form of Payment", , "Error"
            ElseIf AmountTendered = Total Then
                MsgBox "Have A Nice Day!", , "Thank You!"
        End If
           
        
End Sub

Private Sub cmdCheck_Click()
        'select form of payment
        PointOfSale.Visible = True
        TransactionDetails.Visible = False
        Check.Visible = True
        Status.Visible = False
        
End Sub

Private Sub cmdDebitCredit_Click()
        'select form of payment
        PointOfSale.Visible = True
        TransactionDetails.Visible = False
        DebitCredit.Visible = True
        Status.Visible = False
End Sub

Private Sub cmdGiftCard_Click()
        'select form of payment
        PointOfSale.Visible = True
        TransactionDetails.Visible = False
        GiftCard.Visible = True
        Status.Visible = False
End Sub



'RetailPOSandInventoryControl program; TransactionDetails form
'this code was written on Wednesday, November 1, 2006
'edited on Thursday, November 2, 2006
'written by Mark Collette
'the purpose of this form is to select a customer payment method and complete a transaction
'the subroutines consisted of command buttons to select payment methods and complete transactions
