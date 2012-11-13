VERSION 5.00
Begin VB.Form DebitCredit 
   BackColor       =   &H80000010&
   Caption         =   "Customer Details"
   ClientHeight    =   6285
   ClientLeft      =   4620
   ClientTop       =   3060
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   6120
   Begin VB.TextBox txtThree 
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox txtTwo 
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtOne 
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmdProcess 
      BackColor       =   &H008080FF&
      Caption         =   "Process Transaction"
      Height          =   615
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      Width           =   2415
   End
   Begin VB.OptionButton Option4 
      Caption         =   "American Express"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Discover"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      Caption         =   "MasterCard"
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Visa"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblExpirDate 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Enter Expiration Date:"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label lblCardNum 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Enter Card Number:"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblSelect 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Select Card Type"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "DebitCredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdProcess_Click()
        
        If txtOne.Text = "" Then
            MsgBox "Please Enter All Required Info", , "Error"
        ElseIf txtTwo.Text = "" Then
            MsgBox "Please Enter All Required Info", , "Error"
        ElseIf txtThree.Text = "" Then
            MsgBox "Please Enter All Required Info", , "Error"
        Else
            MsgBox "Your Payment Has Been Processed.  Thank You!", , "Thank You"
        DebitCredit.Visible = False
        TransactionDetails.Visible = False
        PointOfSale.Visible = True
        Status.Visible = False
        End If
                
End Sub



'RetailPOSandInventoryControl program; DebitCredit form
'this code was written on Tuesday, October 31, 2006
'this code was edited on Wednesday, November 1, 2006
'this code was revised on Thursday, November 2, 2006
'written by Mark Collette
'the purpose of this form is to take in customer info via Debit or Credit card to process a transaction
'the subroutines took in the info, and displayed a message box to note the info had been received and thank the customer for their business
