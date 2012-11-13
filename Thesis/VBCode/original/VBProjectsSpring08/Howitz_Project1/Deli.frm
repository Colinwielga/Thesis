VERSION 5.00
Begin VB.Form frmDeli 
   BackColor       =   &H0080FF80&
   Caption         =   "Form1"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7335
   ForeColor       =   &H0080FF80&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton cmdWrap 
      Caption         =   "Wrap"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdhard 
      Caption         =   "Hard Taco"
      Height          =   855
      Left            =   2040
      TabIndex        =   8
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdsoft 
      Caption         =   "Soft Taco"
      Height          =   855
      Left            =   2040
      TabIndex        =   7
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cdmnacho 
      Caption         =   "Nachos"
      Height          =   855
      Left            =   2040
      TabIndex        =   6
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdfull 
      Caption         =   "Full Sub"
      Height          =   855
      Left            =   2040
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdthird 
      Caption         =   "1/3 Sub"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmd225 
      Caption         =   "$2.25 Sandwich"
      Height          =   855
      Left            =   2040
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmd200 
      Caption         =   "$2.00 Sandwich"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdhalf 
      Caption         =   "1/2 Sub"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000014&
      Height          =   7575
      Left            =   3840
      ScaleHeight     =   7515
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmDeli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sexton Cash Register
'Form Name:  frmDeli
'Louis Howitz
'March 31, 2008
'All of the deli options are available in this menu.  Items and
'prices will load from an array and print in the picture box.

Private Sub cdmnacho_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Deli(7)
        CartPrices(Items) = DeliPrice(7)
        picResults.Print Deli(7); Tab(20); FormatCurrency(DeliPrice(7))
        
    Close #1
End Sub

Private Sub cmd200_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Deli(3)
        CartPrices(Items) = DeliPrice(3)
        picResults.Print Deli(3); Tab(20); FormatCurrency(DeliPrice(3))
        
    Close #1
End Sub

Private Sub cmd225_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Deli(6)
        CartPrices(Items) = DeliPrice(6)
        picResults.Print Deli(6); Tab(20); FormatCurrency(DeliPrice(6))
        
    Close #1
End Sub

Private Sub cmdBack_Click()
    frmDeli.Hide
    frmTill.Show
    
End Sub

Private Sub cmdfull_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Deli(5)
        CartPrices(Items) = DeliPrice(5)
        picResults.Print Deli(5); Tab(20); FormatCurrency(DeliPrice(5))
        
    Close #1
End Sub

Private Sub cmdhalf_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Deli(1)
        CartPrices(Items) = DeliPrice(1)
        picResults.Print Deli(1); Tab(20); FormatCurrency(DeliPrice(1))
        
    Close #1
End Sub

Private Sub cmdhard_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Deli(8)
        CartPrices(Items) = DeliPrice(8)
        picResults.Print Deli(8); Tab(20); FormatCurrency(DeliPrice(8))
        
    Close #1
End Sub

Private Sub cmdsoft_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Deli(9)
        CartPrices(Items) = DeliPrice(9)
        picResults.Print Deli(9); Tab(20); FormatCurrency(DeliPrice(9))
        
    Close #1
End Sub

Private Sub cmdthird_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Deli(2)
        CartPrices(Items) = DeliPrice(2)
        picResults.Print Deli(2); Tab(20); FormatCurrency(DeliPrice(2))
        
    Close #1
End Sub

Private Sub cmdWrap_Click()

    Items = Items + 1
        ShoppingCart(Items) = Deli(4)
        CartPrices(Items) = DeliPrice(4)
        picResults.Print Deli(4); Tab(20); FormatCurrency(DeliPrice(4))
        
    Close #1
End Sub

Private Sub Form_Load()
'The file is loaded as the form is entered.

Dim CTR As Integer
CTR = 0
    
    Open App.Path & "\Deli.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Deli(CTR), DeliPrice(CTR)
    Loop
    Close #1
End Sub
