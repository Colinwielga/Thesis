VERSION 5.00
Begin VB.Form frmSnack 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8955
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9945
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   7320
      Width           =   2295
   End
   Begin VB.CommandButton cmdjerky 
      Caption         =   "Beef Jerky"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmd225 
      Caption         =   "$2.25 Chips"
      Height          =   855
      Left            =   2160
      TabIndex        =   8
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdChips 
      Caption         =   "$3.49 Chips"
      Height          =   855
      Left            =   2160
      TabIndex        =   7
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmd75 
      Caption         =   "$.75 Chips"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmd25 
      Caption         =   "$.25 Candy"
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdGum 
      Caption         =   "Gum"
      Height          =   855
      Left            =   2160
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdHostess 
      Caption         =   "Hostess Cake"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdOrbit 
      Caption         =   "Orbit Gum"
      Height          =   855
      Left            =   2160
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdCandy 
      Caption         =   "Candy Bar"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   8055
      Left            =   4200
      ScaleHeight     =   7995
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "frmSnack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sexton Cash Register
'Form Name:  frmSnack
'Louis Howitz
'March 31, 2008
'Each snack item is displayed on the buttons.  They will
'print the item and price in the picture box.

Private Sub cmd225_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Snack(9)
        CartPrices(Items) = SnackPrice(9)
        picResults.Print Snack(9); Tab(20); FormatCurrency(SnackPrice(9))
        
    Close #1
End Sub

Private Sub cmd25_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Snack(3)
        CartPrices(Items) = SnackPrice(3)
        picResults.Print Snack(3); Tab(20); FormatCurrency(SnackPrice(3))
        
    Close #1
End Sub

Private Sub cmd75_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Snack(4)
        CartPrices(Items) = SnackPrice(4)
        picResults.Print Snack(4); Tab(20); FormatCurrency(SnackPrice(4))
        
    Close #1
End Sub

Private Sub cmdBack_Click()
    frmSnack.Hide
    frmTill.Show
    
End Sub

Private Sub cmdCandy_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Snack(1)
        CartPrices(Items) = SnackPrice(1)
        picResults.Print Snack(1); Tab(20); FormatCurrency(SnackPrice(1))
        
    Close #1
End Sub

Private Sub cmdChips_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Snack(8)
        CartPrices(Items) = SnackPrice(8)
        picResults.Print Snack(8); Tab(20); FormatCurrency(SnackPrice(8))
        
    Close #1
End Sub

Private Sub cmdGum_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Snack(6)
        CartPrices(Items) = SnackPrice(6)
        picResults.Print Snack(6); Tab(20); FormatCurrency(SnackPrice(6))
        
    Close #1
End Sub

Private Sub cmdHostess_Click()
    Items = Items + 1
        ShoppingCart(Items) = Snack(2)
        CartPrices(Items) = SnackPrice(2)
        picResults.Print Snack(2); Tab(20); FormatCurrency(SnackPrice(2))
        
    Close #1
End Sub

Private Sub cmdjerky_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Snack(5)
        CartPrices(Items) = SnackPrice(5)
        picResults.Print Snack(5); Tab(20); FormatCurrency(SnackPrice(5))
        
    Close #1
End Sub

Private Sub cmdOrbit_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Snack(7)
        CartPrices(Items) = SnackPrice(7)
        picResults.Print Snack(7); Tab(20); FormatCurrency(SnackPrice(7))
        
    Close #1
End Sub

Private Sub Form_Load()
'The file is loaded from an array.

Dim CTR As Integer
CTR = 0
    
    Open App.Path & "\Snack.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Snack(CTR), SnackPrice(CTR)
    Loop
    Close #1
End Sub
