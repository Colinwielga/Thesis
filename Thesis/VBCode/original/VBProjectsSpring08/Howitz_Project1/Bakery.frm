VERSION 5.00
Begin VB.Form frmBakery 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9240
   FillColor       =   &H000080FF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   Picture         =   "Bakery.frx":0000
   ScaleHeight     =   9690
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   8400
      Width           =   2415
   End
   Begin VB.CommandButton cmdPie 
      Caption         =   "Pie"
      Height          =   855
      Left            =   2040
      TabIndex        =   6
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdDonut 
      Caption         =   "Donut"
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdMuffin 
      Caption         =   "Muffin"
      Height          =   855
      Left            =   2040
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdScone 
      Caption         =   "Scone"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdRoll 
      Caption         =   "Roll"
      Height          =   855
      Left            =   2040
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdBagel 
      Caption         =   "Bagel"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000E&
      Height          =   8175
      Left            =   5040
      ScaleHeight     =   8115
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   4995
      Left            =   0
      Picture         =   "Bakery.frx":50AA2
      Top             =   3360
      Width           =   4950
   End
End
Attribute VB_Name = "frmBakery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Project Name: Sexton Cash Register
'Form Name:  frmBakery
'Louis Howitz
'March 31, 2008
'This form includes all of the items available in the bakery
'menu.  The names of the items and prices will print on the
'picture box as the command buttons are pressed.

Private Sub cmdBack_Click()
    frmBakery.Hide
    frmTill.Show
End Sub
Private Sub cmdBagel_Click()
'Each item is hard coded with their line in the array
     
     Items = Items + 1
        ShoppingCart(Items) = BakeFood(1)
        CartPrices(Items) = BakePrice(1)
        picResults.Print BakeFood(1); Tab(20); FormatCurrency(BakePrice(1))
        
    Close #1
End Sub

Private Sub cmdDonut_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = BakeFood(3)
        CartPrices(Items) = BakePrice(3)
        picResults.Print BakeFood(3); Tab(20); FormatCurrency(BakePrice(3))
        
    Close #1
End Sub

Private Sub cmdMuffin_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = BakeFood(4)
        CartPrices(Items) = BakePrice(4)
        picResults.Print BakeFood(4); Tab(20); FormatCurrency(BakePrice(4))
        
    Close #1
End Sub

Private Sub cmdPie_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = BakeFood(6)
        CartPrices(Items) = BakePrice(6)
        picResults.Print BakeFood(6); Tab(20); FormatCurrency(BakePrice(6))
        
    Close #1
End Sub

Private Sub cmdRoll_Click()
    
   Items = Items + 1
        ShoppingCart(Items) = BakeFood(5)
        CartPrices(Items) = BakePrice(5)
        picResults.Print BakeFood(5); Tab(20); FormatCurrency(BakePrice(5))
        
    Close #1
End Sub

Private Sub cmdScone_Click()

    Items = Items + 1
        ShoppingCart(Items) = BakeFood(2)
        CartPrices(Items) = BakePrice(2)
        picResults.Print BakeFood(2); Tab(20); FormatCurrency(BakePrice(2))
        
    Close #1
End Sub

Private Sub Form_Load()
'The array will be loaded as the form opens

Dim CTR As Integer
CTR = 0
    
    Open App.Path & "\Bakery.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, BakeFood(CTR), BakePrice(CTR)
    Loop
    Close #1
End Sub
