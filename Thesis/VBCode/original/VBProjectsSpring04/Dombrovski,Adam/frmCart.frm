VERSION 5.00
Begin VB.Form frmCart 
   BackColor       =   &H80000009&
   Caption         =   "Shopping Cart"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBill 
      BackColor       =   &H00FF8080&
      Height          =   4455
      Left            =   360
      ScaleHeight     =   4395
      ScaleWidth      =   8115
      TabIndex        =   4
      Top             =   1800
      Width           =   8175
   End
   Begin VB.CommandButton cmdViewBill 
      Caption         =   "Click to view Detailed Bill"
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Leave Store"
      Height          =   855
      Left            =   7080
      TabIndex        =   2
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmdBWeight 
      Caption         =   "Go Back to Weight Management Page"
      Height          =   855
      Left            =   2160
      TabIndex        =   1
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmdBSports 
      Caption         =   "Go Back to Athletic Performance Page"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1530
      Left            =   3480
      Picture         =   "frmCart.frx":0000
      Top             =   0
      Width           =   2160
   End
End
Attribute VB_Name = "frmCart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : AdvoCare Store (Dombrovski,Adam.vbp)
'Form Name : frmCart(Shopping Cart)
'Author: Adam Dombrovski
'Date Written: March 15, 2004
'Purpose: This form shows a print out of all the products the user
    'has put in their shopping cart, how much they spent on each item
    'and the total cost.  It also allows the user to go back to either
    'the Athletic Performance page or the weight management page.  Also
    'ends the program

Private Sub cmdViewBill_Click()
picBill.Cls
picBill.Print Tab(25); "Here is your detailed bill."
picBill.Print "---------------------------------------------------------------------------------------------------------------"
picBill.Print "Name of Product"; Tab(55); "Total per product"
picBill.Print
picBill.Print "*****************************************************************************************"
For l = 1 To ctr
    picBill.Print Cart(l); Tab(55); FormatCurrency(CartPrice(l))
Next l
picBill.Print Tab(55); "-------------"
picBill.Print Tab(40); "Total:"; Tab(55); FormatCurrency(runningTotal)
End Sub


Private Sub cmdBSports_Click()
frmCart.Hide
frmSports.Show
End Sub

Private Sub cmdBWeight_Click()
frmCart.Hide
frmWeight.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub



Private Sub Form_Load()
cmdViewBill.Enabled = True
End Sub
