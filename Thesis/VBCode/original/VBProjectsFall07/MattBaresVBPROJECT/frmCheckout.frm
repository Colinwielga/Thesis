VERSION 5.00
Begin VB.Form frmCheckout 
   Caption         =   "Target"
   ClientHeight    =   9975
   ClientLeft      =   5205
   ClientTop       =   2670
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   Picture         =   "frmCheckout.frx":0000
   ScaleHeight     =   9975
   ScaleWidth      =   7215
   Begin VB.Frame fraCheckout 
      Caption         =   "Welcome to the Checkout Aisle"
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   3000
      Width           =   7215
      Begin VB.CheckBox chkCoupon 
         Caption         =   "If you were given a coupon, check the box to the left."
         Enabled         =   0   'False
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         Top             =   360
         Width           =   3015
      End
      Begin VB.PictureBox picDog 
         Height          =   1935
         Left            =   4080
         Picture         =   "frmCheckout.frx":10EE1
         ScaleHeight     =   1875
         ScaleWidth      =   1875
         TabIndex        =   6
         Top             =   4200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdPay 
         Caption         =   "Pay and Leave"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3120
         TabIndex        =   5
         Top             =   3240
         Width           =   3855
      End
      Begin VB.PictureBox picTotal 
         Height          =   1455
         Left            =   3240
         ScaleHeight     =   1395
         ScaleWidth      =   3555
         TabIndex        =   4
         Top             =   1680
         Width           =   3615
      End
      Begin VB.PictureBox picCart 
         Height          =   6135
         Left            =   120
         ScaleHeight     =   6075
         ScaleWidth      =   2835
         TabIndex        =   3
         Top             =   720
         Width           =   2895
      End
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "Display my cart"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdCheckout 
         Caption         =   "Checkout"
         Height          =   615
         Left            =   3120
         TabIndex        =   1
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label lblThanks 
         Caption         =   "Thanks for shopping at Target!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   7
         Top             =   6240
         Visible         =   0   'False
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmCheckout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Subtotal As Single, CouponCheck As Integer, Coupon As String, Total As Single, TotalwCoupon As Single

Private Sub chkCoupon_Click()
'This subroutine checks to see if a coupon was given to the user for having 10 or more
'items in their shopping cart(As seen at the end of cdmDisplay)
Total = Subtotal
'Here I used a checkbox which has not been covered in class

If chkCoupon.Value = 1 Then
    Coupon = InputBox("Please type in your coupon as it was given to you.", "Target")
    'The following is a nested If statement that applies the coupon discount of 5%
    'to the user's total price.
    If Coupon = "TARGET" Then
        TotalwCoupon = (Total * 0.95)
    Else
        MsgBox "That is an invalid coupon. Please try again.", , "Target"
        Coupon = InputBox("Please type in your coupon as it was given to you.", "Target")
        TotalwCoupon = (Total * 0.95)
    End If
End If

End Sub

Private Sub cmdCheckOut_Click()
'This subroutine prints the subtotal and total in a picture box, which can be affected by the use of a coupon
'It also enables the "Pay and Leave" button
cmdPay.Enabled = True
cmdDisplay.Enabled = False

Total = Subtotal

picTotal.Cls
picTotal.Print "Your subtotal is:"
picTotal.Print Tab(37); FormatCurrency(Subtotal)

If Coupon = "TARGET" Then
    picTotal.Print "plus tax (6.5%) and 5% off brings your total to:"
    picTotal.Print Tab(37); FormatCurrency(TotalwCoupon + TotalwCoupon * 0.065)
Else
    picTotal.Print "plus tax (6.5%) brings your total to:"
    picTotal.Print Tab(37); FormatCurrency(Total + Total * 0.065)
End If

End Sub

Private Sub cmdDisplay_Click()
Dim MasterInventory(0 To 500) As Integer, MasterName(0 To 500) As String, MasterPrice(0 To 500) As Single
Dim InventoryC(0 To 100) As Integer
Dim InventoryK(0 To 100) As Integer
Dim InventoryT(0 To 100) As Integer
Dim InventoryE(0 To 100) As Integer
Dim InventoryF(0 To 100) As Integer
Dim CTR As Integer, CTR2 As Integer, CTR3 As Integer, pos As Integer, X As Integer, CTR4 As Integer, CTR5 As Integer, CTR6 As Integer

'This subroutine loads the file for each cart into a corresponding array for comparison to the 'master list' array.

'Once the display button has been clicked, it cannot be clicked again.
cmdDisplay.Enabled = False

Open App.Path & "\MasterList.txt" For Input As #1
    
CTR = 0

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, MasterInventory(CTR), MasterName(CTR), MasterPrice(CTR)
Loop
Close #1

Open App.Path & "\ClothingCart.txt" For Input As #2

CTR2 = 0

Do While Not EOF(2)
    CTR2 = CTR2 + 1
    Input #2, InventoryC(CTR2)
Loop
Close #2

Open App.Path & "\FurnitureCart.txt" For Input As #3

CTR3 = 0

Do While Not EOF(3)
    CTR3 = CTR3 + 1
    Input #3, InventoryF(CTR3)
Loop
Close #3

Open App.Path & "\KitchenCart.txt" For Input As #4

CTR4 = 0

Do While Not EOF(4)
    CTR4 = CTR4 + 1
    Input #4, InventoryK(CTR4)
Loop
Close #4

Open App.Path & "\ToysCart.txt" For Input As #5

CTR5 = 0

Do While Not EOF(5)
    CTR5 = CTR5 + 1
    Input #5, InventoryT(CTR2)
Loop
Close #5

Open App.Path & "\ElectronicsCart.txt" For Input As #6

CTR6 = 0

Do While Not EOF(6)
    CTR6 = CTR6 + 1
    Input #6, InventoryE(CTR6)
Loop
Close #6


picCart.Cls
picCart.Print "#"; Tab(6); "Name"; Tab(31); "Price"
picCart.Print "*********************************************"

'this section compares each departments array containing inventory numbers, to the master list array.
'when the program finds a match it looks up the corresponding name and price to the inventory number and displays
'it in a picture box.

'for each item purchased, the program adds 1 to 'CouponCheck' in order to determine if a coupon will be given.
'for a coupon to be given, 'CouponCheck' must be >= 10, i.e. ten items bought.

For pos = 1 To CTR2
    For X = 1 To CTR
        If MasterInventory(X) = InventoryC(pos) Then
            picCart.Print MasterInventory(X); Tab(6); MasterName(X); Tab(31); FormatCurrency(MasterPrice(X))
            Subtotal = Subtotal + MasterPrice(X)
            CouponCheck = CouponCheck + 1
        End If
    Next X
Next pos

For pos = 1 To CTR3
    For X = 1 To CTR
        If MasterInventory(X) = InventoryF(pos) Then
            picCart.Print MasterInventory(X); Tab(6); MasterName(X); Tab(31); FormatCurrency(MasterPrice(X))
            Subtotal = Subtotal + MasterPrice(X)
            CouponCheck = CouponCheck + 1
        End If
    Next X
Next pos

For pos = 1 To CTR4
    For X = 1 To CTR
        If MasterInventory(X) = InventoryK(pos) Then
            picCart.Print MasterInventory(X); Tab(6); MasterName(X); Tab(31); FormatCurrency(MasterPrice(X))
            Subtotal = Subtotal + MasterPrice(X)
            CouponCheck = CouponCheck + 1
        End If
    Next X
Next pos

For pos = 1 To CTR5
    For X = 1 To CTR
        If MasterInventory(X) = InventoryT(pos) Then
            picCart.Print MasterInventory(X); Tab(6); MasterName(X); Tab(31); FormatCurrency(MasterPrice(X))
            Subtotal = Subtotal + MasterPrice(X)
            CouponCheck = CouponCheck + 1
        End If
    Next X
Next pos

For pos = 1 To CTR6
    For X = 1 To CTR
        If MasterInventory(X) = InventoryE(pos) Then
            picCart.Print MasterInventory(X); Tab(6); MasterName(X); Tab(31); FormatCurrency(MasterPrice(X))
            Subtotal = Subtotal + MasterPrice(X)
            CouponCheck = CouponCheck + 1
        End If
    Next X
Next pos

'this section determines if 'CouponCheck' is greater than or equal to ten, and if so, provides the user with a coupon
'and enables the checkbox for using and entering a coupon.

If CouponCheck >= 10 Then
    picCart.Print ""
    picCart.Print "Your coupon is: TARGET"
    chkCoupon.Enabled = True
End If

End Sub

Private Sub cmdPay_Click()
'this subroutine pretends that the user has made a money transaction and thanks them  by displaying a picture box
'containing the target dog from a file, and a label saying 'Thanks for shopping at Target'.
'It also displays a msgbox thanking the user, which when closed, hides the checkout form and shows the entrance form.

lblThanks.Visible = True
picDog.Visible = True
MsgBox "Thanks for shopping at Target! We hope you enjoy your purchases.", , "Target"
frmCheckout.Hide
frmEntrance.Show
End Sub
