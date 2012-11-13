VERSION 5.00
Begin VB.Form frmCheckout 
   Caption         =   "Ben's Hockey Goods"
   ClientHeight    =   12300
   ClientLeft      =   6585
   ClientTop       =   1350
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   ScaleHeight     =   12300
   ScaleWidth      =   13350
   Begin VB.Frame fraCheckout 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome to the Checkout Aisle"
      Height          =   12255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13335
      Begin VB.CommandButton cmdCheckout 
         BackColor       =   &H000000FF&
         Caption         =   "Checkout"
         BeginProperty Font 
            Name            =   "Goudy Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   11400
         Width           =   3975
      End
      Begin VB.CommandButton cmdDisplay 
         BackColor       =   &H0000FF00&
         Caption         =   "Display my cart"
         BeginProperty Font 
            Name            =   "Goudy Old Style"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   3615
      End
      Begin VB.PictureBox picCart 
         BackColor       =   &H00FFFFFF&
         Height          =   7335
         Left            =   9000
         ScaleHeight     =   7275
         ScaleWidth      =   4155
         TabIndex        =   4
         Top             =   1080
         Width           =   4215
      End
      Begin VB.PictureBox picTotal 
         BackColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   5760
         ScaleHeight     =   1395
         ScaleWidth      =   3195
         TabIndex        =   3
         Top             =   4200
         Width           =   3255
      End
      Begin VB.CommandButton cmdPay 
         BackColor       =   &H000000FF&
         Caption         =   "Pay and Leave"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   8520
         Width           =   3975
      End
      Begin VB.CheckBox chkCoupon 
         BackColor       =   &H0000FF00&
         Caption         =   "If you were given a coupon, check the box to the left."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Goudy Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6120
         TabIndex        =   1
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Image Imageplayer 
         Height          =   12405
         Left            =   0
         Picture         =   "frmCheckout.frx":0000
         Top             =   0
         Width           =   9000
      End
      Begin VB.Label lblThanks 
         BackColor       =   &H0000FF00&
         Caption         =   "Thanks for shopping at Ben's Hockey Goods!"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   9000
         TabIndex        =   7
         Top             =   9720
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
'Ben's Hockey Store
'frmCheckout
'Ben Bartelt
'3/26/08
'The purupose of this form is to display all the items that the user has purchased from the previous departments.
'It doesn't let the user pay and leave until they checkout. This form also counts the number of items purchased.
'If the user has purchased 10 or more items from my store they are given a 5% discount on their overall order.
'It also has a picture of all required hockey gear that is necessary to be a protected hockey player.
'The form also calculates a 6.5% sales tax to the subtotal.
'I have comments under most subroutines to describe what each button is doing.
Option Explicit
Dim Subtotal As Single, CouponCheck As Integer, Coupon As String, Total As Single, TotalwCoupon As Single

Private Sub chkCoupon_Click()
'This subroutine checks to see if a coupon was given to the user for having 10 or more
'items in their shopping cart
Total = Subtotal
'Check box is next used

If chkCoupon.Value = 1 Then
    Coupon = InputBox("Please type in your coupon exactly as shown.", "Ben's Hockey Goods")
    'The following is a nested If statement that applies the coupon discount of 5%
    'to the user's total price.
    If Coupon = "Ben's Hockey Goods" Then
        TotalwCoupon = (Total * 0.95)
    Else
        MsgBox "That is an invalid coupon. Please try again.", , "Ben's Hockey Goods"
        Coupon = InputBox("Please type in your coupon exactly as shown.", "Ben's Hockey Goods")
        TotalwCoupon = (Total * 0.95)
    End If
End If

End Sub

Private Sub cmdCheckOut_Click()
'This subroutine prints the subtotal and total price may be affected if coupon is enable. It also enables pay and leave
cmdPay.Enabled = True
cmdDisplay.Enabled = False

Total = Subtotal

picTotal.Cls
picTotal.Print "Your subtotal is:"
picTotal.Print Tab(37); FormatCurrency(Subtotal)

If Coupon = "Ben's Hockey Goods" Then
    picTotal.Print "plus tax (6.5%) and 5% off brings your total to:"
    picTotal.Print Tab(37); FormatCurrency(TotalwCoupon + TotalwCoupon * 0.065)
Else
    picTotal.Print "plus tax (6.5%) brings your total to:"
    picTotal.Print Tab(37); FormatCurrency(Total + Total * 0.065)
End If

End Sub

Private Sub cmdDisplay_Click()
Dim MasterInventory(0 To 500) As Integer, MasterName(0 To 500) As String, MasterPrice(0 To 500) As Single
Dim InventoryP(0 To 100) As Integer
Dim InventoryH(0 To 100) As Integer
Dim InventoryK(0 To 100) As Integer
Dim InventoryA(0 To 100) As Integer
Dim InventoryS(0 To 100) As Integer
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

Open App.Path & "\Paddingcart.txt" For Input As #2

CTR2 = 0

Do While Not EOF(2)
    CTR2 = CTR2 + 1
    Input #2, InventoryP(CTR2)
Loop
Close #2

Open App.Path & "\Stickscart.txt" For Input As #3

CTR3 = 0

Do While Not EOF(3)
    CTR3 = CTR3 + 1
    Input #3, InventoryS(CTR3)
Loop
Close #3

Open App.Path & "\Helmetscart.txt" For Input As #4

CTR4 = 0

Do While Not EOF(4)
    CTR4 = CTR4 + 1
    Input #4, InventoryH(CTR4)
Loop
Close #4

Open App.Path & "\Skatescart.txt" For Input As #5

CTR5 = 0

Do While Not EOF(5)
    CTR5 = CTR5 + 1
    Input #5, InventoryK(CTR2)
Loop
Close #5

Open App.Path & "\accessoriescart.txt" For Input As #6

CTR6 = 0

Do While Not EOF(6)
    CTR6 = CTR6 + 1
    Input #6, InventoryA(CTR6)
Loop
Close #6


picCart.Cls
picCart.Print "#"; Tab(6); "Name"; Tab(31); "Price"
picCart.Print "****************************************************************************"

'this section compares each departments array containing inventory numbers, to the master list array.
'when the program finds a match it looks up the corresponding name and price to the inventory number and displays
'it in a picture box.
'for each item purchased, the program adds 1 to 'CouponCheck' in order to determine if a coupon will be given.
For pos = 1 To CTR2
    For X = 1 To CTR
        If MasterInventory(X) = InventoryP(pos) Then
            picCart.Print MasterInventory(X); Tab(6); MasterName(X); Tab(31); FormatCurrency(MasterPrice(X))
            Subtotal = Subtotal + MasterPrice(X)
            CouponCheck = CouponCheck + 1
        End If
    Next X
Next pos

For pos = 1 To CTR3
    For X = 1 To CTR
        If MasterInventory(X) = InventoryS(pos) Then
            picCart.Print MasterInventory(X); Tab(6); MasterName(X); Tab(31); FormatCurrency(MasterPrice(X))
            Subtotal = Subtotal + MasterPrice(X)
            CouponCheck = CouponCheck + 1
        End If
    Next X
Next pos

For pos = 1 To CTR4
    For X = 1 To CTR
        If MasterInventory(X) = InventoryH(pos) Then
            picCart.Print MasterInventory(X); Tab(6); MasterName(X); Tab(31); FormatCurrency(MasterPrice(X))
            Subtotal = Subtotal + MasterPrice(X)
            CouponCheck = CouponCheck + 1
        End If
    Next X
Next pos

For pos = 1 To CTR5
    For X = 1 To CTR
        If MasterInventory(X) = InventoryK(pos) Then
            picCart.Print MasterInventory(X); Tab(6); MasterName(X); Tab(31); FormatCurrency(MasterPrice(X))
            Subtotal = Subtotal + MasterPrice(X)
            CouponCheck = CouponCheck + 1
        End If
    Next X
Next pos

For pos = 1 To CTR6
    For X = 1 To CTR
        If MasterInventory(X) = InventoryA(pos) Then
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
    picCart.Print "Your coupon is: Ben's Hockey Goods"
    chkCoupon.Enabled = True
End If

End Sub

Private Sub cmdPay_Click()
'this subroutine thanks the user for their purches  by displaying a picture box
lblThanks.Visible = True
MsgBox "Thanks for shopping at Ben's Hockey Goods! We hope you enjoy your equipment.", , "Ben's Hockey Goods"
frmCheckout.Hide
frmEntrance.Show
End Sub

Private Sub Imageplayer_Click()
' just an image of a hockey player with all the required equipment to play hockey without getting hurt.
End Sub
