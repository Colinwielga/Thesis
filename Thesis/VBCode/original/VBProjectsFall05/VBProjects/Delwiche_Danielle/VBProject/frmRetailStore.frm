VERSION 5.00
Begin VB.Form frmDaniRetailStore 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Dani's Retail Store"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOther 
      Height          =   3855
      Left            =   6840
      ScaleHeight     =   3795
      ScaleWidth      =   3075
      TabIndex        =   22
      Top             =   3600
      Width           =   3135
   End
   Begin VB.CommandButton cmdBegin 
      Caption         =   "Begin Shopping"
      Height          =   375
      Left            =   8040
      TabIndex        =   21
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search For An Item By Price"
      Height          =   855
      Left            =   3960
      TabIndex        =   20
      Top             =   5040
      Width           =   2415
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit Store"
      Height          =   855
      Left            =   3960
      TabIndex        =   19
      Top             =   6360
      Width           =   2415
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Find Total"
      Height          =   855
      Left            =   3960
      TabIndex        =   18
      Top             =   3720
      Width           =   2415
   End
   Begin VB.PictureBox picOutput 
      Height          =   4215
      Left            =   240
      ScaleHeight     =   4155
      ScaleWidth      =   3195
      TabIndex        =   17
      Top             =   3240
      Width           =   3255
   End
   Begin VB.CommandButton cmdTee 
      Caption         =   "Add T-Shirt"
      Height          =   495
      Left            =   7680
      TabIndex        =   16
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdBaseball 
      Caption         =   "Add Baseball Tee"
      Height          =   495
      Left            =   5160
      TabIndex        =   15
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdTank 
      Caption         =   "Add Tank Top"
      Height          =   495
      Left            =   2640
      TabIndex        =   14
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdShorts 
      Caption         =   "Add Shorts"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdPants 
      Caption         =   "Add Pants"
      Height          =   495
      Left            =   8400
      TabIndex        =   12
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdNecklace 
      Caption         =   "Add Necklace"
      Height          =   495
      Left            =   6000
      TabIndex        =   11
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdJeans 
      Caption         =   "Add Jeans"
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdCoat 
      BackColor       =   &H80000009&
      Caption         =   "Add Coat"
      Height          =   495
      Left            =   960
      TabIndex        =   9
      Top             =   720
      Width           =   1455
   End
   Begin VB.PictureBox picTee 
      BackColor       =   &H80000009&
      Height          =   855
      Left            =   9240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmRetailStore.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   7
      Top             =   2040
      Width           =   855
   End
   Begin VB.PictureBox picBaseball 
      BackColor       =   &H80000009&
      Height          =   975
      Left            =   6720
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmRetailStore.frx":03F1
      ScaleHeight     =   915
      ScaleWidth      =   795
      TabIndex        =   6
      Top             =   2040
      Width           =   855
   End
   Begin VB.PictureBox picTank 
      BackColor       =   &H80000009&
      Height          =   1215
      Left            =   4200
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmRetailStore.frx":086B
      ScaleHeight     =   1155
      ScaleWidth      =   795
      TabIndex        =   5
      Top             =   2040
      Width           =   855
   End
   Begin VB.PictureBox picShorts 
      BackColor       =   &H80000009&
      Height          =   975
      Left            =   1680
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmRetailStore.frx":0D5C
      ScaleHeight     =   915
      ScaleWidth      =   795
      TabIndex        =   4
      Top             =   2040
      Width           =   855
   End
   Begin VB.PictureBox picPants 
      BackColor       =   &H80000009&
      Height          =   1215
      Left            =   7560
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmRetailStore.frx":118A
      ScaleHeight     =   1155
      ScaleWidth      =   675
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.PictureBox picNecklace 
      BackColor       =   &H80000009&
      Height          =   1095
      Left            =   5160
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmRetailStore.frx":15C6
      ScaleHeight     =   1035
      ScaleWidth      =   675
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.PictureBox picJeans 
      BackColor       =   &H80000009&
      Height          =   975
      Left            =   2520
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmRetailStore.frx":1D93
      ScaleHeight     =   915
      ScaleWidth      =   795
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.PictureBox picCoat 
      BackColor       =   &H80000009&
      FillColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmRetailStore.frx":2342
      ScaleHeight     =   795
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblClick 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Click On the Picture For Product Information, And Click on the Add Buttons to Purchase:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "frmDaniRetailStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dani's Retail Store
'by Danielle Delwiche
'Completed Monday, October 24, 2005
'This program is designed to allow a "customer" to purchase items, see a
    'running subtotal, and final total, and also search for items by price.
    
Option Explicit
    Dim I As Integer
    Dim Sum As Single
    Dim Product(1 To 8) As String
    Dim Price(1 To 8) As Single
    Dim NotFound As Boolean
    
'This sub adds an item to the shopper's cart, and displays the subtotal
Private Sub cmdBaseball_Click()
    Dim Baseball As String
    Baseball = "Baseball"
    NotFound = True
    I = 0
    Do While NotFound And I <= 8
        I = I + 1
        If Baseball = Product(I) Then
            NotFound = False
            picOutput.Print Product(I), FormatCurrency(Price(I))
            Sum = Sum + Price(I)
            picOther.Cls
            picOther.Print "Your current Subtotal is:", FormatCurrency(Sum)
        End If
    Loop
End Sub

'This sub adds an item to the shopper's cart, and displays the subtotal
Private Sub cmdCoat_Click()
    Dim Coat As String
    Coat = "Coat"
    NotFound = True
    I = 0
    Do While NotFound And I <= 8
        I = I + 1
        If Coat = Product(I) Then
            NotFound = False
            picOutput.Print Product(I), FormatCurrency(Price(I))
            Sum = Sum + Price(I)
            picOther.Cls
            picOther.Print "Your current Subtotal is:", FormatCurrency(Sum)
        End If
    Loop
End Sub

'This command ends the program
Private Sub cmdExit_Click()
    End
End Sub

'This command sets up the array in which all the products and prices are listed
'This allows us to search through the products by price or name throughout the program
Private Sub cmdBegin_Click()
    Sum = 0
    I = 0
    Open App.Path & "\Apparel.txt" For Input As #1
    Do Until EOF(1)
        I = I + 1
        Input #1, Product(I), Price(I)
    Loop
End Sub

'This sub adds an item to the shopper's cart, and displays the subtotal
Private Sub cmdJeans_Click()
    Dim Jeans As String
    Jeans = "Jeans"
    NotFound = True
    I = 0
    Do While NotFound And I <= 8
        I = I + 1
        If Jeans = Product(I) Then
            NotFound = False
            picOutput.Print Product(I), FormatCurrency(Price(I))
            Sum = Sum + Price(I)
            picOther.Cls
            picOther.Print "Your current Subtotal is:", FormatCurrency(Sum)
        End If
    Loop
End Sub

'This sub adds an item to the shopper's cart, and displays the subtotal
Private Sub cmdNecklace_Click()
    Dim Necklace As String
    Necklace = "Necklace"
    NotFound = True
    I = 0
    Do While NotFound And I <= 8
        I = I + 1
        If Necklace = Product(I) Then
            NotFound = False
            picOutput.Print Product(I), FormatCurrency(Price(I))
            Sum = Sum + Price(I)
            picOther.Cls
            picOther.Print "Your current Subtotal is:", FormatCurrency(Sum)
        End If
    Loop
End Sub

'This sub adds an item to the shopper's cart, and displays the subtotal
Private Sub cmdPants_Click()
    Dim Pants As String
    Pants = "Pants"
    NotFound = True
    I = 0
    Do While NotFound And I <= 8
        I = I + 1
        If Pants = Product(I) Then
            NotFound = False
            picOutput.Print Product(I), FormatCurrency(Price(I))
            Sum = Sum + Price(I)
            picOther.Cls
            picOther.Print "Your current Subtotal is:", FormatCurrency(Sum)
        End If
    Loop
End Sub

'This command allows the shopper to search through the items by maximum price and provides a readout of items matching the search
Private Sub cmdSearch_Click()
    Dim Search As Single
    picOther.Cls
    Search = InputBox("Enter the highest price you wish to search for:", "Price Search")
    picOther.Print "Items Matching Your Search:"
    For I = 1 To 8
        If Search >= Price(I) Then
            picOther.Print Product(I), FormatCurrency(Price(I))
        End If
    Next I
End Sub

'This sub adds an item to the shopper's cart, and displays the subtotal
Private Sub cmdShorts_Click()
    Dim Shorts As String
    Shorts = "Shorts"
    NotFound = True
    I = 0
    Do While NotFound And I <= 8
        I = I + 1
        If Shorts = Product(I) Then
            NotFound = False
            picOutput.Print Product(I), FormatCurrency(Price(I))
            Sum = Sum + Price(I)
            picOther.Cls
            picOther.Print "Your current Subtotal is:", FormatCurrency(Sum)
        End If
    Loop
End Sub

'This sub adds an item to the shopper's cart, and displays the subtotal
Private Sub cmdTank_Click()
    Dim Tank As String
    Tank = "Tank"
    NotFound = True
    I = 0
    Do While NotFound And I <= 8
        I = I + 1
        If Tank = Product(I) Then
            NotFound = False
            picOutput.Print Product(I), FormatCurrency(Price(I))
            Sum = Sum + Price(I)
            picOther.Cls
            picOther.Print "Your current Subtotal is:", FormatCurrency(Sum)
        End If
    Loop
End Sub

'This sub adds an item to the shopper's cart, and displays the subtotal
Private Sub cmdTee_Click()
    Dim Tee As String
    Tee = "T-Shirt"
    NotFound = True
    I = 0
    Do While NotFound And I <= 8
        I = I + 1
        If Tee = Product(I) Then
            NotFound = False
            picOutput.Print Product(I), FormatCurrency(Price(I))
            Sum = Sum + Price(I)
            picOther.Cls
            picOther.Print "Your current Subtotal is:", FormatCurrency(Sum)
        End If
    Loop
End Sub

'This sub provides a final total for the shopper, including subtotal, tax, and grand total
'This sub is set up so that a shopper can continue to add items to this running total even after the grand total is given
Private Sub cmdTotal_Click()
    Dim Tax As Single
    Dim Total As Single
    Tax = Sum * 0.065
    Total = Sum + Tax
    picOther.Cls
    picOther.Print "Subtotal:", FormatCurrency(Sum)
    picOther.Print "Tax:", FormatCurrency(Tax)
    picOther.Print "******************************"
    picOther.Print "Total:", FormatCurrency(Total)
End Sub

'This sub provides a pop-up with a product description
Private Sub picBaseball_Click()
    MsgBox "This Baseball Tee is perfect for a casual day, or your favorite game!  Priced at $25.", , "Baseball Tee"
End Sub

'This sub provides a pop-up with a product description
Private Sub picJeans_Click()
    MsgBox "These jeans are made of the finest cotton blend, and stonewashed for that visual perfection!  Priced at $45.", , "Jeans"
End Sub

'This sub provides a pop-up with a product description
Private Sub picCoat_Click()
    MsgBox "This pea coat is made of a wool and cotton blend, made to keep you very warm in the winter cold!  Priced at $175.", , "Coat"
End Sub

'This sub provides a pop-up with a product description
Private Sub picNecklace_Click()
    MsgBox "This necklace is hand-made, with clay beads.  Priced at $15.", , "Necklace"
End Sub

'This sub provides a pop-up with a product description
Private Sub picPants_Click()
    MsgBox "These pants are made of a cotton and spandex blend, adding a hint of stretch for an amazing fit! Priced at $48.", , "Pants"
End Sub

'This sub provides a pop-up with a product description
Private Sub picShorts_Click()
    MsgBox "A variation of our pants, these short are a comfy cotton blend.  Priced at $35.", , "Shorts"
End Sub

'This sub provides a pop-up with a product description
Private Sub picTank_Click()
    MsgBox "This tank top is 100% cotton to give you breathing room for summer days, and for layering in the fall!  Priced at $20.", , "Tank Top"
End Sub

'This sub provides a pop-up with a product description
Private Sub picTee_Click()
    MsgBox "Dress it up, Dress it down!  Never has a 100% cotton T-Shirt been so versatile!  Priced at $22.", , "T-Shirt"
End Sub
