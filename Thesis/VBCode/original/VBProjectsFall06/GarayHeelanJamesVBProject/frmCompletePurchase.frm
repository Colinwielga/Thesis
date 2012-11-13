VERSION 5.00
Begin VB.Form frmCompletePurchase 
   BackColor       =   &H000040C0&
   Caption         =   "Complete Purchase"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLogOut 
      Caption         =   "Log Out"
      Height          =   855
      Left            =   600
      TabIndex        =   2
      Top             =   6120
      Width           =   3015
   End
   Begin VB.PictureBox picResults 
      Height          =   3255
      Left            =   4680
      ScaleHeight     =   3195
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   4080
      Width           =   5055
   End
   Begin VB.CommandButton cmdPurchase 
      Caption         =   "Complete Transaction!"
      Height          =   1215
      Left            =   5520
      TabIndex        =   0
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   4695
      Left            =   120
      Picture         =   "frmCompletePurchase.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   4485
   End
End
Attribute VB_Name = "frmCompletePurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'James Delivers, CSCI 130 Visual Basic Project
'frmCompletePurchase
'Written by James Garay Heelan
'on 11-2-06
'This form is where the user may purchase his or her groceries for delivery.

Option Explicit

Private Sub cmdLogOut_Click()
    End 'exits program
End Sub

Private Sub cmdPurchase_Click()
Dim Total As Double

    Total = (Sum + Sum * 0.15 + Sum * 0.065) 'defines the total cost to the user as the sum of the products purchased, plus the shipping and handling costs, plus the sales tax
    
    picResults.Print , " "; FormatCurrency(Sum) 'displays the purchase amount
    picResults.Print , "+"; FormatCurrency(Sum * 0.15); " Shipping and Handling" 'displays in the picturebox the shipping and handling costs
    picResults.Print , "+"; FormatCurrency(Sum * 0.065); " Tax" 'prints in the picturebox the sales tax on the goods
    picResults.Print , "__________________" 'prints a line
    picResults.Print "Total, "; " ", FormatCurrency(Total) 'prints the total cost to the user
    picResults.Print FormatCurrency(Total) & " will be charged to your " & PaymentMethod(PurchaseCode) & " number " & CreditCardNumber(PurchaseCode) 'verifies that the total will be charged via the payment method entered by the user at registration the total purchase cost
    picResults.Print "Your groceries will arrive at the below address within 48 hours."
    picResults.Print " "
    picResults.Print Names(PurchaseCode) 'Prints the user's address, as entered in his or her registration
    picResults.Print Address(PurchaseCode)
    picResults.Print City(PurchaseCode), State(PurchaseCode)
    picResults.Print Zip(PurchaseCode)
    picResults.Print " "
    picResults.Print "Thank you for shopping James Delivers!"
    picResults.Print "We hope to deliver to you again, soon, "; Names(PurchaseCode) 'Thanks the user by name
End Sub
