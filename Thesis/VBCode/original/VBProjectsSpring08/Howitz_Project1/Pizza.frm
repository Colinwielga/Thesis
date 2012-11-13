VERSION 5.00
Begin VB.Form frmPizza 
   BackColor       =   &H000000FF&
   Caption         =   "Form2"
   ClientHeight    =   9975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   9975
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   1095
      Left            =   240
      TabIndex        =   8
      Top             =   8280
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000E&
      Height          =   6975
      Left            =   5280
      ScaleHeight     =   6915
      ScaleWidth      =   4515
      TabIndex        =   7
      Top             =   2400
      Width           =   4575
   End
   Begin VB.CommandButton cmdGarlic 
      Caption         =   "Garlic Bread"
      Height          =   1095
      Left            =   2880
      TabIndex        =   5
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton cmdExtra 
      Caption         =   "Extra Toppings"
      Height          =   1095
      Left            =   7680
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton cdmBread 
      Caption         =   "Bread Stick"
      Height          =   1095
      Left            =   2880
      TabIndex        =   3
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton cmdSauce 
      Caption         =   "Extra Sauce"
      Height          =   1095
      Left            =   5280
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton cmdWhole 
      Caption         =   "Whole Pizza"
      Height          =   1095
      Left            =   480
      TabIndex        =   1
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton cdmSlice 
      Caption         =   "Pizza by the slice"
      Height          =   1095
      Left            =   480
      MaskColor       =   &H00004000&
      TabIndex        =   0
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   3465
      Left            =   240
      Picture         =   "Pizza.frx":0000
      Top             =   3720
      Width           =   5025
   End
   Begin VB.Label lblPizza 
      BackColor       =   &H0000FFFF&
      Caption         =   "PIZZA"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   6
      Top             =   120
      Width           =   9375
   End
End
Attribute VB_Name = "frmPizza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Project Name: Sexton Cash Register
'Form Name:  frmPizza
'Louis Howitz
'March 31, 2008
'This form displays all of the pizza options at Sexton.
'The buttons print the item and price in the picture box.

Private Sub cdmBread_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Pizza(2)
        CartPrices(Items) = PizzaPrice(2)
        picResults.Print Pizza(2); Tab(20); FormatCurrency(PizzaPrice(2))
        
    Close #1
End Sub

Private Sub cdmSlice_Click()

    Items = Items + 1
        ShoppingCart(Items) = Pizza(1)
        CartPrices(Items) = PizzaPrice(1)
        picResults.Print Pizza(1); Tab(20); FormatCurrency(PizzaPrice(1))
        
    Close #1
End Sub

Private Sub cmdBack_Click()
    frmPizza.Hide
    frmTill.Show
    
End Sub

Private Sub cmdExtra_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Pizza(3)
        CartPrices(Items) = PizzaPrice(3)
        picResults.Print Pizza(3); Tab(20); FormatCurrency(PizzaPrice(3))
        
    Close #1
End Sub

Private Sub cmdGarlic_Click()

    Items = Items + 1
        ShoppingCart(Items) = Pizza(6)
        CartPrices(Items) = PizzaPrice(6)
        picResults.Print Pizza(6); Tab(20); FormatCurrency(PizzaPrice(6))
        
    Close #1
End Sub

Private Sub cmdSauce_Click()

    Items = Items + 1
        ShoppingCart(Items) = Pizza(3)
        CartPrices(Items) = PizzaPrice(3)
        picResults.Print Pizza(3); Tab(20); FormatCurrency(PizzaPrice(3))
        
    Close #1
End Sub

Private Sub cmdWhole_Click()

    Items = Items + 1
        ShoppingCart(Items) = Pizza(5)
        CartPrices(Items) = PizzaPrice(5)
        picResults.Print Pizza(5); Tab(20); FormatCurrency(PizzaPrice(5))
        
    Close #1
End Sub

Private Sub Form_Load()
'The file is loaded from an array.

Dim CTR As Integer
CTR = 0
    
    Open App.Path & "\Pizza.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Pizza(CTR), PizzaPrice(CTR)
    Loop
    Close #1
End Sub
