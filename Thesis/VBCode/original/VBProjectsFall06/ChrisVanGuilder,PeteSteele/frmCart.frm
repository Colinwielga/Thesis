VERSION 5.00
Begin VB.Form frmCart 
   BackColor       =   &H00800000&
   Caption         =   "Your Shopping Cart!"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9645
   FillColor       =   &H00800000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtZip 
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox txtState 
      Height          =   495
      Left            =   2280
      TabIndex        =   11
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtAddress 
      Height          =   615
      Left            =   2280
      TabIndex        =   10
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtLast 
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtFirst 
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturnHome 
      Caption         =   "Return to the Home Page"
      Height          =   735
      Left            =   960
      TabIndex        =   3
      Top             =   5280
      Width           =   2655
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Your Cart"
      Height          =   855
      Left            =   2400
      TabIndex        =   2
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total Your Purchases"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   4200
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   5055
      Left            =   5160
      ScaleHeight     =   4995
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label lblZip 
      Caption         =   "Enter Your Zip Code"
      Height          =   495
      Left            =   360
      TabIndex        =   13
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label lblState 
      Caption         =   "State"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblAddress 
      Caption         =   "Enter Your Street Address"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblLastName 
      Caption         =   "Enter Your Last Name"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblFirstName 
      Caption         =   "Enter Your First Name"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmCart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MLBonline.vbp
'frmCart.frm
'Chris Van Guilder and Pete Steele, 11/2/2006
'this form is the "checkout" for the user. Here, this form displays the Merchandise and Gear totals, along with a grand total.
'If the purchase is over $200, there is a $15 discount subtracted from total.
'This form asks for the buyers information for shipping purposes.

Option Explicit

Private Sub cmdClear_Click() 'clears cart and sets all sums to zero
    picResults.Cls
    Sumall = 0
    GearSum = 0
    MerchandiseSum = 0
End Sub

Private Sub cmdReturnHome_Click() 'returns to the home page
    frmCart.Visible = False
    frmHomepage.Visible = True
End Sub

Private Sub cmdTotal_Click() 'computes total from Merchandise Purchases and Gear Purchases and subtracts discount
    picResults.Print "**************************************************"
    Sumall = MerchandiseSum + GearSum
    picResults.Print "Your Merchandise Total is : ", FormatCurrency(MerchandiseSum, 2) 'Merchandise total
    picResults.Print "Your Gear Total is : ", FormatCurrency(GearSum, 2) 'Gear total
    If Sumall >= 200 Then
        picResults.Print "Your Total is: ", , FormatCurrency(Sumall, 2)  'decides whether or not discount is applicable
        picResults.Print "You spent over $200!!!   ", "     -15"
        Sumall = Sumall - 15
        picResults.Print "Your Grand Total is ", FormatCurrency(Sumall, 2)
    Else
        picResults.Print "Your Grand Total is ", FormatCurrency(Sumall, 2)
    End If
    picResults.Print "Thank you for your total order of " & FormatCurrency(Sumall); " " 'prints receipt and confirms delivery information"
    picResults.Print "Thank You " & txtFirst; " " & txtLast
    picResults.Print "Your order has been sent to: "
    picResults.Print txtAddress; ", "; txtState; " "; txtZip
    picResults.Print "Your Receipt # is: 196534"
    picResults.Print "Use this Receipt # to Enter the Drawing!"
End Sub

