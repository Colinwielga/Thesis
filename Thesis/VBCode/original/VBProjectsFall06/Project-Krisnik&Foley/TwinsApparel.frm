VERSION 5.00
Begin VB.Form TwinsApparel 
   BackColor       =   &H0080C0FF&
   Caption         =   "Twins Apparel"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   5640
      TabIndex        =   17
      Top             =   6360
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   8280
      TabIndex        =   15
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5640
      TabIndex        =   13
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H0080FFFF&
      Caption         =   "Clear Order"
      Height          =   615
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton cmdBaseball 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Purchase Baseball"
      Height          =   375
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdPurchaseSweat 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Purchase Sweatshirt"
      Height          =   375
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdFill 
      BackColor       =   &H0080FFFF&
      Caption         =   "Shop"
      Height          =   615
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton cmdCheckout 
      BackColor       =   &H0080FFFF&
      Caption         =   "Checkout"
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6000
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000E&
      Height          =   4455
      Left            =   600
      ScaleHeight     =   4395
      ScaleWidth      =   3915
      TabIndex        =   7
      Top             =   1440
      Width           =   3975
   End
   Begin VB.PictureBox Picture5 
      Height          =   1215
      Left            =   8520
      Picture         =   "TwinsApparel.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.PictureBox Picture4 
      Height          =   1215
      Left            =   8520
      Picture         =   "TwinsApparel.frx":148E
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      Height          =   1215
      Left            =   8280
      Picture         =   "TwinsApparel.frx":1F5F
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   3
      Top             =   5760
      Width           =   1215
   End
   Begin VB.PictureBox pictshirt 
      Height          =   1695
      Left            =   5760
      Picture         =   "TwinsApparel.frx":3138
      ScaleHeight     =   1635
      ScaleWidth      =   1395
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   5520
      Picture         =   "TwinsApparel.frx":4A9F
      ScaleHeight     =   1635
      ScaleWidth      =   1875
      TabIndex        =   1
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Return To Homepage"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7200
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Enter Quantity of Hats:"
      Height          =   255
      Left            =   5640
      TabIndex        =   18
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Enter Quantity of T-Shirts:"
      Height          =   255
      Left            =   8280
      TabIndex        =   16
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Enter Quantity of Jerseys:"
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lblShop 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Click On The Shop Button To Get Started!"
      BeginProperty Font 
         Name            =   "Eras Bold ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "TwinsApparel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Project Name: Twins Baseball
' Form Name: Twins Apparel
' Authors: Jake Krisnik & Mike Foley
' Date Written: October 23, 2006
' Form Objective: To provide the user with the ability to view some common Twins apparel
'                 and souveniers with the opportunity to add them to a shopping cart and
'                 view their total purchase price after they have added the items they
'                 desire with a 5% shipping charge computed to the total. They can also
'                 navigate away from the store to return to the Twins Homepage.
Option Explicit
Dim RunningTotal As Single

Private Sub cmdBaseball_Click()
' This button allows the user to click on it to add a baseball to the shopping cart.
' It also keeps track of a running total which is used to calculate in the checkout
' stage of the shopping process.
    Dim Baseball As Single
    Baseball = 8.95
    RunningTotal = RunningTotal + Baseball
    picResults.Print "Baseball", "                              "; FormatCurrency(Baseball)
End Sub

Private Sub cmdCheckout_Click()
' This command button is the one that calculates the running total of all the items the user has
' added to their shopping cart. It then calculates the shipping charges applied to the purchase
' and shows the total. It then gives the user a thanks for shopping for Minnesota Twins apparel.
    Dim Total As Single, Shipping As Single
    MsgBox "Be advised that there is a 5% shipping charge for your order"
    Shipping = RunningTotal * 0.05
    picResults.Print , "                              "; "----------------"
    picResults.Print "Subtotal: ", "                              "; FormatCurrency(RunningTotal, 2)
    picResults.Print "Shipping:", "                              "; FormatCurrency(Shipping, 2)
    picResults.Print , "                              "; "----------------"
    picResults.Print "Total:", "                              "; FormatCurrency(RunningTotal + Shipping, 2)
    MsgBox "Thank you for your order!"
End Sub

Private Sub cmdClear_Click()
' This button allows the user to clear the output box and everything in it to start their
' shopping over.
    picResults.Cls
    RunningTotal = 0
End Sub

Private Sub cmdFill_Click()
' This button fills the array of items and their corresponding prices into the output box to
' allow the user to view them before clicking on the items to add them to the cart. It then
' offers a message prompt the the user to click on the items next to add them to the cart.
    Dim I As Integer
    Dim Item(1 To 5) As String
    Dim Price(1 To 5) As Single
        picResults.Print "Item                                               Price"
        picResults.Print "_____________________________________"
    Open App.Path & "\Shop.txt" For Input As #1
    For I = 1 To 5
        Input #1, Item(I), Price(I)
        picResults.Print Item(I); "              ", "       "; FormatCurrency(Price(I), 2)
    Next I
        picResults.Print "********************************************************"
    Close 1
    MsgBox "Click on the available items to add them to your cart and select the checkout button when finished"
End Sub



Private Sub cmdPurchaseSweat_Click()
' This button allows the user to click on it to add a sweatshirt to the shopping cart.
' It also keeps track of a running total which is used to calculate in the checkout
' stage of the shopping process.
    Dim Sweatshirt As Single
    Sweatshirt = 75.95
    RunningTotal = RunningTotal + Sweatshirt
    picResults.Print "Sweatshirt", "                              "; FormatCurrency(Sweatshirt)
End Sub


Private Sub cmdReturn_Click()
' This command button allows the user to navigate away from the Twins Apparel form and
' return to the Homepage.
    HomePage.Show
    TwinsApparel.Hide
End Sub


Private Sub Text1_Change()
' This button allows the user to add a desired amount of jerseys into the text box and
' ultimately shopping cart. It also keeps track of a running total which is used to calculate in the checkout
' calculate in the checkout stage of the shopping process.
    Dim Jersey As Single
    Dim X As Integer
    X = Text1.Text
    Jersey = 99.95 * X
    RunningTotal = RunningTotal + Jersey
    picResults.Print "Jersey", "                              "; FormatCurrency(Jersey)

End Sub

Private Sub Text2_Change()
' This button allows the user to add a desired amount of t-shirts into the text box and
' ultimately shopping cart. It also keeps track of a running total which is used to calculate in the checkout
' calculate in the checkout stage of the shopping process.
    Dim Tshirt As Single
    Dim X As Integer
    X = Text2.Text
    Tshirt = 25.95 * X
    RunningTotal = RunningTotal + Tshirt
    picResults.Print "T-Shirt", "                              "; FormatCurrency(Tshirt)
End Sub

Private Sub Text3_Change()
' This button allows the user to add a desired amount of hats into the text box and
' ultimately shopping cart. It also keeps track of a running total which is used to calculate in the checkout
' calculate in the checkout stage of the shopping process.
    Dim Hat As Single
    Dim X As Integer
    X = Text3.Text
    Hat = 15.95 * X
    RunningTotal = RunningTotal + Hat
    picResults.Print "Twins Hat", "                              "; FormatCurrency(Hat)
End Sub
