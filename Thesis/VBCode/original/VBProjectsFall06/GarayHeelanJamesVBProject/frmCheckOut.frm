VERSION 5.00
Begin VB.Form frmCheckOut 
   Caption         =   "Check Out"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13380
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   13380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdContinueShopping 
      Caption         =   "Continue Shopping"
      Height          =   855
      Left            =   1080
      TabIndex        =   4
      Top             =   6480
      Width           =   3615
   End
   Begin VB.CommandButton cmdLogOut 
      Caption         =   "Log Out"
      Height          =   855
      Left            =   1080
      TabIndex        =   3
      Top             =   9120
      Width           =   3615
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Purchase Contents of Cart"
      Height          =   855
      Left            =   1080
      TabIndex        =   2
      Top             =   7800
      Width           =   3615
   End
   Begin VB.CommandButton cmdCart 
      Caption         =   "View Contents in Your Shopping Cart"
      Height          =   855
      Left            =   1080
      TabIndex        =   1
      Top             =   5160
      Width           =   3615
   End
   Begin VB.PictureBox picResults 
      Height          =   9735
      Left            =   7920
      ScaleHeight     =   9675
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   480
      Width           =   5175
   End
   Begin VB.Image imgCheckOut 
      Height          =   10935
      Left            =   0
      Picture         =   "frmCheckOut.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13800
   End
End
Attribute VB_Name = "frmCheckOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'James Delivers, CSCI 130 Visual Basic Project
'frmCheckOut
'Written by James Garay Heelan
'on 11-2-06
'The CheckOut form allows the user to view the items in his or her shopping cart
'and decide whether or not to purchase the goods or to continue shopping.

Option Explicit

Private Sub cmdBuy_Click()
frmCheckOut.Hide 'hides the checkout form and
frmCompletePurchase.Show 'displays the completepurchase form
End Sub

Private Sub cmdCart_Click()

Dim Item(1 To 100) As String, Price(1 To 100) As String 'loads the array variables
Dim I As Integer, J As Integer 'loads the counting variables
Dim Found As Boolean
I = 0
Found = False

    picResults.Cls 'clears the picturebox of its contents
    picResults.Print "Item, "; "Price" 'prints a header in the picturebox letting the user know what is being printed below it
    
    Open App.Path & "/PurchasedItems.txt" For Input As #5 'the shopping card file is opened for reading
    
    Do Until EOF(5) 'the search is instructed to continue till it reaches the end of the file
        I = I + 1 'the counter is increased by 1, to move to the next grocery item on the list
        Input #5, Item(I), Price(I) 'the file is sorted into two arrays, one for item name and the other for item price
        picResults.Print Item(I), Price(I) 'the currently loaded item and its price are printed in the picturebox
    Loop 'the search loops
    Close #5 'the shopping cart file is closed
    
    picResults.Print " " 'a blank line is printed
    picResults.Print "Grand Total: "; FormatCurrency(Sum) 'the sum of the total product cost, without tax or shipping and handling, is printed int the picturebox

End Sub

Private Sub cmdContinueShopping_Click()
    frmCheckOut.Hide 'hides the checkout form and allows the user to
    frmGroceryStore.Show 'continue shopping by bringing them back to the central menu
End Sub

Private Sub cmdLogOut_Click()
    End 'exits the program
End Sub

