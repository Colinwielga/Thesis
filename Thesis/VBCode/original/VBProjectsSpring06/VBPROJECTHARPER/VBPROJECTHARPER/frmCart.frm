VERSION 5.00
Begin VB.Form frmCart 
   BackColor       =   &H8000000D&
   Caption         =   "My Cart"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClearcart 
      Caption         =   "Clear Cart"
      Height          =   495
      Left            =   6240
      TabIndex        =   15
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   495
      Left            =   4920
      TabIndex        =   14
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdHP 
      Caption         =   "Return to Home Page"
      Height          =   495
      Left            =   1320
      TabIndex        =   13
      Top             =   4560
      Width           =   1815
   End
   Begin VB.ComboBox Method 
      Height          =   315
      ItemData        =   "frmCart.frx":0000
      Left            =   2160
      List            =   "frmCart.frx":000D
      TabIndex        =   12
      Top             =   2760
      Width           =   1935
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00808080&
      Caption         =   "Pay on Delivery"
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   3960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C000&
      Caption         =   "Credit Card"
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox State 
      Height          =   315
      ItemData        =   "frmCart.frx":003D
      Left            =   2160
      List            =   "frmCart.frx":0050
      TabIndex        =   4
      Text            =   "Select State..."
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtaddress 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtlast 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtname 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000E&
      Height          =   4335
      Left            =   4560
      ScaleHeight     =   4275
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label6 
      Caption         =   "By: Ben Harper"
      Height          =   255
      Left            =   7560
      TabIndex        =   16
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Caption         =   "State"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Caption         =   "Method of Delivery"
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "Delivery Address"
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Last Name"
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "First Name"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmCart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Buy Drums Online (OnlineDrums.vbp)
'frmCart (frmCart)
'Ben Harper
'3/23/06
'This Form displays the cost of each item, a description of the item in the cart and a grand total.
'If the purchase is over $200, there is a discount subtracted from total. This form asks for the buyers
'information, payment method, and shipping method.













Private Sub cmdClearcart_Click() 'clears cart and sets all sums to zero
picResults.Cls
Sumall = 0
Cymbalsum = 0
Drumsum = 0
Accsum = 0
End Sub

Private Sub cmdHP_Click()  'returns to HomePage
frmCart.Visible = False
frmHomePage.Visible = True
End Sub

Private Sub cmdTotal_Click()     'computes total from all three sectors and subtracts discount
    picResults.Print "**************************************************"
    Sumall = Drumsum + Cymbalsum
    picResults.Print "Your Drum Total is : ", FormatCurrency(Drumsum, 2) 'drum total
    picResults.Print "Your Cymbal Total is : ", FormatCurrency(Cymbalsum, 2) 'cymbal total
    picResults.Print "Your Accessory Total is : ", FormatCurrency(Accsum, 2) 'accessory total
    If Sumall >= 200 Then
        picResults.Print "Your Total is: ", , FormatCurrency(Sumall, 2)  'decides whether discount is deserved or not
        picResults.Print "You spent over $200!!!   ", "     -10"
        Sumall = Sumall - 10
        picResults.Print "Your Grand Total is ", FormatCurrency(Sumall, 2)
     Else
        picResults.Print "Your Grand Total is ", FormatCurrency(Sumall, 2)
   End If
    Option1.Visible = True    'payment methods become visible after total is computed
    Option2.Visible = True
   End Sub



Private Sub Option1_Click()
X = InputBox("Please enter your 16 digit credit card number", "Card information", "e.g. 4563 7865 9912 9036") 'gathers card information for credit card payment
X = InputBox("Please enter your card's expiration date", "card information", "mm/yy")
MsgBox "Thank You " & txtname, , "Your Order has been placed"
 
End Sub

Private Sub Option2_Click()
    MsgBox "Thank you for your total order of " & FormatCurrency(Sumall), , "Money Order Placed!" 'prints receipt and confirms delivery information
    picResults.Print "Thank You " & txtname & txtlast
    picResults.Print "Your order has been sent to: "
    picResults.Print txtaddress; ", " & State & ", "; Method
End Sub

