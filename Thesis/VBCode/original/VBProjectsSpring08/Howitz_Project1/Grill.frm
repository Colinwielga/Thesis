VERSION 5.00
Begin VB.Form frmGrill 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8295
   FillColor       =   &H80000001&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   12
      Top             =   7200
      Width           =   2535
   End
   Begin VB.CommandButton cmdcheese 
      Caption         =   "Grilled Cheese"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdrings 
      Caption         =   "Onion rings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   10
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdshrimp 
      Caption         =   "Shrimp Basket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdfish 
      Caption         =   "Fish Burger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   8
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdtots 
      Caption         =   "Tots"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdbasket 
      Caption         =   "Chicken Basket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   6
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdfries 
      Caption         =   "Fries"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdburger 
      Caption         =   "Burger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdChicken 
      Caption         =   "Chicken Burger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdgrillchk 
      Caption         =   "Grilled Chicken"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdchsburger 
      Caption         =   "Cheese Burger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   4080
      ScaleHeight     =   7875
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label lblgrill 
      BackColor       =   &H0000FFFF&
      Caption         =   "GRILL"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmGrill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Sexton Cash Register
'Form Name:  frmGrill
'Louis Howitz
'March 31, 2008
'All of the grill options are available on this form.  The buttons
'will print the item and price in the picture box.

Private Sub cmdBack_Click()
    
    frmGrill.Hide
    frmTill.Show
End Sub

Private Sub cmdbasket_Click()
    
    
    Items = Items + 1
        ShoppingCart(Items) = GrillFood(9)
        CartPrices(Items) = GrillPrice(9)
        picResults.Print GrillFood(9); Tab(20); FormatCurrency(GrillPrice(9))
        
    
    Close #1
End Sub

Private Sub cmdburger_Click()

    Items = Items + 1
        ShoppingCart(Items) = GrillFood(2)
        CartPrices(Items) = GrillPrice(2)
        picResults.Print GrillFood(2); Tab(20); FormatCurrency(GrillPrice(2))
        
    Close #1
End Sub

Private Sub cmdcheese_Click()
     Items = Items + 1
        ShoppingCart(Items) = GrillFood(6)
        CartPrices(Items) = GrillPrice(6)
        picResults.Print GrillFood(6); Tab(20); FormatCurrency(GrillPrice(6))
        
    Close #1
End Sub

Private Sub cmdChicken_Click()
   
    Items = Items + 1
        ShoppingCart(Items) = GrillFood(8)
        CartPrices(Items) = GrillPrice(8)
        picResults.Print GrillFood(8); Tab(20); FormatCurrency(GrillPrice(8))
        
End Sub

Private Sub cmdchsburger_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = GrillFood(1)
        CartPrices(Items) = GrillPrice(1)
        picResults.Print GrillFood(1); Tab(20); FormatCurrency(GrillPrice(1))
        
    Close #1
        
End Sub


Private Sub cmdfish_Click()
    
     Items = Items + 1
        ShoppingCart(Items) = GrillFood(10)
        CartPrices(Items) = GrillPrice(10)
        picResults.Print GrillFood(10); Tab(20); FormatCurrency(GrillPrice(10))
        
    Close #1
End Sub

Private Sub cmdfries_Click()
     
     Items = Items + 1
        ShoppingCart(Items) = GrillFood(7)
        CartPrices(Items) = GrillPrice(7)
        picResults.Print GrillFood(7); Tab(20); FormatCurrency(GrillPrice(7))
        
    Close #1
End Sub

Private Sub cmdgrillchk_Click()

     Items = Items + 1
        ShoppingCart(Items) = GrillFood(3)
        CartPrices(Items) = GrillPrice(3)
        picResults.Print GrillFood(3); Tab(20); FormatCurrency(GrillPrice(3))
        
    Close #1
End Sub

Private Sub cmdrings_Click()

     Items = Items + 1
        ShoppingCart(Items) = GrillFood(11)
        CartPrices(Items) = GrillPrice(11)
        picResults.Print GrillFood(11); Tab(20); FormatCurrency(GrillPrice(11))
        
    Close #1
End Sub

Private Sub cmdshrimp_Click()
     
     Items = Items + 1
        ShoppingCart(Items) = GrillFood(5)
        CartPrices(Items) = GrillPrice(5)
        picResults.Print GrillFood(5); Tab(20); FormatCurrency(GrillPrice(5))
        
    Close #1
End Sub

Private Sub cmdtots_Click()
    
     Items = Items + 1
        ShoppingCart(Items) = GrillFood(4)
        CartPrices(Items) = GrillPrice(4)
        picResults.Print GrillFood(4); Tab(20); FormatCurrency(GrillPrice(4))
        
    Close #1
End Sub

Private Sub Form_Load()
'The file is loaded from the array.

Dim CTR As Integer
CTR = 0
    
    Open App.Path & "\Grill.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, GrillFood(CTR), GrillPrice(CTR)
    Loop
    Close #1
    
        
    
    
End Sub
