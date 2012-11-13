VERSION 5.00
Begin VB.Form frmBev 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   10230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9270
   FillColor       =   &H00FF0000&
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   8400
      Width           =   2175
   End
   Begin VB.CommandButton cmdice 
      Caption         =   "Iced Tea"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdh2o 
      Caption         =   "Bottled Water"
      Height          =   855
      Left            =   2040
      TabIndex        =   8
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdEnergy 
      Caption         =   "Energy Drink"
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdfrap 
      Caption         =   "Frappuccino"
      Height          =   855
      Left            =   2040
      TabIndex        =   6
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdcoffee 
      Caption         =   "Coffee"
      Height          =   855
      Left            =   2040
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdfountain 
      Caption         =   "Fountain Soda"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdTea 
      Caption         =   "Tea"
      Height          =   855
      Left            =   2040
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdcap 
      Caption         =   "Cappuccino"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmd20oz 
      Caption         =   "20 oz Soda"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      Height          =   7935
      Left            =   5280
      ScaleHeight     =   7875
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   2505
      Left            =   120
      Picture         =   "Bev.frx":0000
      Top             =   5640
      Width           =   3600
   End
End
Attribute VB_Name = "frmBev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sexton Cash Register
'Form Name:  frmBev
'Louis Howitz
'March 31, 2008
'The form shows the options of beverages available.  Each button
'will print the name of the item and price.

Private Sub cmd20oz_Click()

    Items = Items + 1
        ShoppingCart(Items) = Drink(1)
        CartPrices(Items) = DrinkPrice(1)
        picResults.Print Drink(1); Tab(20); FormatCurrency(DrinkPrice(1))
        
    Close #1
End Sub

Private Sub cmdBack_Click()
    frmBev.Hide
    frmTill.Show
End Sub

Private Sub cmdcap_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Drink(3)
        CartPrices(Items) = DrinkPrice(3)
        picResults.Print Drink(3); Tab(20); FormatCurrency(DrinkPrice(3))
        
    Close #1
End Sub

Private Sub cmdcoffee_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Drink(6)
        CartPrices(Items) = DrinkPrice(6)
        picResults.Print Drink(6); Tab(20); FormatCurrency(DrinkPrice(6))
        
    Close #1
End Sub

Private Sub cmdEnergy_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Drink(4)
        CartPrices(Items) = DrinkPrice(4)
        picResults.Print Drink(4); Tab(20); FormatCurrency(DrinkPrice(4))
        
    Close #1
End Sub

Private Sub cmdfountain_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Drink(2)
        CartPrices(Items) = DrinkPrice(2)
        picResults.Print Drink(2); Tab(20); FormatCurrency(DrinkPrice(2))
        
    Close #1
End Sub

Private Sub cmdfrap_Click()

    Items = Items + 1
        ShoppingCart(Items) = Drink(8)
        CartPrices(Items) = DrinkPrice(8)
        picResults.Print Drink(8); Tab(20); FormatCurrency(DrinkPrice(8))
        
    Close #1
End Sub

Private Sub cmdh2o_Click()

    Items = Items + 1
        ShoppingCart(Items) = Drink(9)
        CartPrices(Items) = DrinkPrice(9)
        picResults.Print Drink(9); Tab(20); FormatCurrency(DrinkPrice(9))
        
    Close #1
End Sub

Private Sub cmdice_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Drink(5)
        CartPrices(Items) = DrinkPrice(5)
        picResults.Print Drink(5); Tab(20); FormatCurrency(DrinkPrice(5))
        
    Close #1
End Sub

Private Sub cmdTea_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Drink(7)
        CartPrices(Items) = DrinkPrice(7)
        picResults.Print Drink(7); Tab(20); FormatCurrency(DrinkPrice(7))
        
    Close #1
End Sub

Private Sub Form_Load()

Dim CTR As Integer
CTR = 0
    
    Open App.Path & "\Drinks.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Drink(CTR), DrinkPrice(CTR)
    Loop
    Close #1
End Sub
