VERSION 5.00
Begin VB.Form frmMerchandise 
   BackColor       =   &H00000000&
   Caption         =   "2006 Championship Merchandise"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   Picture         =   "frmMerchandise.frx":0000
   ScaleHeight     =   7350
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Compute Total Merchandise Purchase"
      Height          =   855
      Left            =   480
      TabIndex        =   7
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton cmdCart 
      Caption         =   "Add Merchandise Total to My Cart"
      Height          =   1095
      Left            =   480
      TabIndex        =   6
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Total"
      Height          =   615
      Left            =   3840
      TabIndex        =   5
      Top             =   6240
      Width           =   2295
   End
   Begin VB.PictureBox picResultsMerchandise 
      Height          =   1575
      Left            =   3240
      ScaleHeight     =   1515
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   4200
      Width           =   3375
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return To Home Page"
      Height          =   1215
      Left            =   7080
      TabIndex        =   3
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label lblShirt 
      BackColor       =   &H00000080&
      Caption         =   "2006 World Series T-Shirt $25.00"
      Height          =   855
      Left            =   7080
      TabIndex        =   2
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label lblSweatshirt 
      BackColor       =   &H00000080&
      Caption         =   "2006 World Series Sweatshirt $50.00"
      Height          =   855
      Left            =   4200
      TabIndex        =   1
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label lblHat 
      BackColor       =   &H00000080&
      Caption         =   "2006 World Series Hat $15.00"
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Image imgSweatshirt 
      Height          =   2820
      Left            =   3960
      Picture         =   "frmMerchandise.frx":F985
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2460
   End
   Begin VB.Image imgTShirt 
      Height          =   2820
      Left            =   6960
      Picture         =   "frmMerchandise.frx":134E6
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2580
   End
   Begin VB.Image imgHat 
      Height          =   3300
      Left            =   240
      Picture         =   "frmMerchandise.frx":15FD0
      Top             =   360
      Width           =   3300
   End
End
Attribute VB_Name = "frmMerchandise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MLBonline.vbp
'frmMerchandise.frm
'Chris Van Guilder and Pete Steele, 11/2/2006
'this form displays the price of select 2006 World Series Championship merchandise.
'The user clicks on the images to add the item to the potential total.
'this form exports the merchandise total to the shopping cart.

Option Explicit
Private Sub cmdCart_Click()
    frmCart.Visible = True                 'sends merchandise total to cart
    frmMerchandise.Visible = False
    frmCart.picResults.Print "Merchandise Purchases", "Total"
    frmCart.picResults.Print "*************************************************"
    frmCart.picResults.Print "You Merchandise Total is: ", FormatCurrency(MerchandiseSum)
End Sub
Private Sub cmdClear_Click() 'clears total Merchandise cost
    picResultsMerchandise.Cls
    MerchandiseSum = 0
    cmdCart.Visible = False
End Sub

Private Sub cmdCompute_Click() 'computes total Merchandise cost
    picResultsMerchandise.Print "*************************************"
    picResultsMerchandise.Print "Your Merchandise Total is: ", FormatCurrency(MerchandiseSum)
    cmdCart.Visible = True
End Sub

Private Sub cmdReturn_Click() 'returns user to home page
    frmMerchandise.Visible = False
    frmHomepage.Visible = True
End Sub

Private Sub imgHat_Click() 'Adds hat cost to total bill and prints the purchase
    picResultsMerchandise.Print "World Series Hat ", FormatCurrency(15, 2)
    MerchandiseSum = MerchandiseSum + 15
End Sub

Private Sub imgSweatshirt_Click() 'add sweatshirt cost to total bill and prints the purchase
    picResultsMerchandise.Print "World Series Sweatshirt ", FormatCurrency(50, 2)
    MerchandiseSum = MerchandiseSum + 50
End Sub


Private Sub imgTShirt_Click() 'add t-shirt cost to total and prints the purchase
    picResultsMerchandise.Print "World Series T-Shirt ", FormatCurrency(25, 2)
    MerchandiseSum = MerchandiseSum + 25
End Sub

