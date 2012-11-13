VERSION 5.00
Begin VB.Form frmProductSearch 
   BackColor       =   &H000040C0&
   Caption         =   "The Campground!"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse on your own."
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   5280
      Width           =   4095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add this item to cart."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   6
      Top             =   3720
      Width           =   4095
   End
   Begin VB.CommandButton cmdGoToCheckout 
      Caption         =   "Proceed to checkout."
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      TabIndex        =   5
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Leave the store with no purchases."
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2760
      TabIndex        =   4
      Top             =   6720
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      Height          =   5055
      Left            =   4920
      ScaleHeight     =   4995
      ScaleWidth      =   4275
      TabIndex        =   3
      Top             =   2640
      Width           =   4335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search for a specific product."
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   2640
      Width           =   4095
   End
   Begin VB.PictureBox picWendy2 
      Height          =   2055
      Left            =   7080
      Picture         =   "frmProductSearch.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblHelpYou 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "How can I help you today?"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "frmProductSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Product As String
Dim CTR As Integer

Private Sub cmdAdd_Click()
'adds the specified item to the item subtotal/quantities item using an If statement
'stores the overall subtotal for use in the Checkout form
'lets user know that their item has been added to their cart
SleepingBag = 74.99
Tent = 199.99
RegJacket = 124.99
XXLJacket = 134.99
MessKit = 14.99
If Product = "Sleeping Bag" Then
    SBCTR = SBCTR + 1
    SleepingBagSub = SleepingBag * SBCTR
ElseIf Product = "Tent" Then
    TCTR = TCTR + 1
    TentSub = Tent * TCTR
ElseIf Product = "S Jacket" Or Product = "M Jacket" Or Product = "L Jacket" Or Product = "XL Jacket" Then
    RJCTR = RJCTR + 1
    JacketSub = RegJacket * RJCTR + XXLJacket * XJCTR
ElseIf Product = "XXL Jacket" Then
    XJCTR = XJCTR + 1
    JacketSub = RegJacket * RJCTR + XXLJacket * XJCTR
ElseIf Product = "Mess Kit" Then
    MKCTR = MKCTR + 1
    MessKitSub = MessKit * MKCTR
End If
picResults.Print "Item successfully added to cart."
Subtotal = SleepingBagSub + TentSub + JacketSub + MessKitSub
End Sub

Private Sub cmdBrowse_Click()
'takes the user to the Browse form using the Visible property
frmProductSearch.Hide
frmBrowse.Show
End Sub

Private Sub cmdGoToCheckout_Click()
'takes the user to the Checkout form using the Visible property
frmProductSearch.Hide
frmCheckout.Show
End Sub


Private Sub cmdQuit_Click()
'ends the program
End
End Sub

Private Sub cmdSearch_Click()
'loads the array of products and prices using a Do Loop, then uses a match and stop
'loop to search for a product and its price
'prints product name and price if found and gives an error if not found
'assigns Product in place of Name(Pos) for use in other subroutines
Dim Pos As Integer
Dim Name(1 To 100) As String
Dim Price(1 To 100) As Single
Dim Found As Boolean
Dim SProduct As String
picResults.Cls
cmdAdd.Enabled = False
CTR = 0
Open App.Path & "\Products.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Name(CTR), Price(CTR)
Loop
Close #1
SProduct = InputBox("Search for a product. If you are looking for a jacket, type the size (S, M, L, XL, or XXL) followed by the word 'Jacket'.", "Product Search")
Found = False And Pos = 0
Do While Found = False And Pos < CTR
    Pos = Pos + 1
    If LCase(SProduct) = LCase(Name(Pos)) Then
        Found = True
    End If
Loop
If Found = True Then
    picResults.Print "Product"; Tab(40); "Price"
    picResults.Print "****************************************************************"
    picResults.Print Name(Pos); Tab(40); FormatCurrency(Price(Pos))
    cmdAdd.Enabled = True
Else
    picResults.Print "Sorry, we don't carry that product."
End If
Product = Name(Pos)
End Sub
