VERSION 5.00
Begin VB.Form frmSoup 
   BackColor       =   &H00000080&
   Caption         =   "Form1"
   ClientHeight    =   9795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   ForeColor       =   &H00000080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9795
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Back"
      Height          =   975
      Left            =   120
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8400
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdSpecial 
      Caption         =   "Special Soup"
      Height          =   855
      Left            =   2400
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdBread 
      Caption         =   "With Bread"
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdCup 
      Caption         =   "Cup-O-Soup"
      Height          =   855
      Left            =   2400
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdBowl 
      Caption         =   "Bowl-O-Soup"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   8295
      Left            =   4560
      ScaleHeight     =   8235
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   6210
      Left            =   0
      Picture         =   "Soup.frx":0000
      Top             =   2520
      Width           =   4500
   End
End
Attribute VB_Name = "frmSoup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sexton Cash Register
'Form Name:  frmSoup
'Louis Howitz
'March 31, 2008
'All of the soup items are displayed on the buttons.  Each button
'will print the item and price in the picture box.

Private Sub cmdBack_Click()
    frmSoup.Hide
    frmTill.Show
    
End Sub


Private Sub cmdBowl_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Soup(1)
        CartPrices(Items) = SoupPrice(1)
        picResults.Print Soup(1); Tab(20); FormatCurrency(SoupPrice(1))
        
    Close #1
End Sub

Private Sub cmdBread_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Soup(2)
        CartPrices(Items) = SoupPrice(2)
        picResults.Print Soup(2); Tab(20); FormatCurrency(SoupPrice(2))
        
    Close #1
    
End Sub

Private Sub cmdCup_Click()

    Items = Items + 1
        ShoppingCart(Items) = Soup(3)
        CartPrices(Items) = SoupPrice(3)
        picResults.Print Soup(3); Tab(20); FormatCurrency(SoupPrice(3))
        
    Close #1
End Sub

Private Sub cmdSpecial_Click()

    Items = Items + 1
        ShoppingCart(Items) = Soup(4)
        CartPrices(Items) = SoupPrice(4)
        picResults.Print Soup(4); Tab(20); FormatCurrency(SoupPrice(4))
        
    Close #1
End Sub

Private Sub Form_Load()
    'The file is loaded from an array.
    
    Dim CTR As Integer
    CTR = 0
    
    Open App.Path & "\Soup.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Soup(CTR), SoupPrice(CTR)
    Loop
    Close #1
End Sub
