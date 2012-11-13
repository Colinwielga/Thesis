VERSION 5.00
Begin VB.Form frmPay 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cdmSources 
      Caption         =   "See Source form"
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort by Price"
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdPay 
      Caption         =   "Pay (Quit)"
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   7440
      Width           =   2415
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000E&
      Height          =   8175
      Left            =   3240
      ScaleHeight     =   8115
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "frmPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sexton Cash Register
'Form Name:  frmPay
'Louis Howitz
'March 31, 2008
'This is the form that will print all of the items and prices from
'each different form.  All the food buttons pressed will be displayed
'as well as the total price.

Dim Total As Single

Private Sub cdmSources_Click()
    
    frmPay.Hide
    frmSources.Show
    
End Sub

Private Sub cmdBack_Click()
 
 frmPay.Hide
 frmTill.Show
 
End Sub

Private Sub cmdPay_Click()
    End
End Sub

Private Sub cmdSort_Click()
'This is a failed attempt at the bubble sort.  I wanted to
'sort the CartPrices from largest to smallest to demonstrate the sort.
'It would also be useful if people wanted to compare prices when
'shopping at Sexton.  To see what is too expensive or a good deal.
'Not working properly.

    Dim Pass As Integer, Pos As Integer, CTR As Integer, Temp As Integer
    
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If CartPrices(Pos) > CartPrices(Pos + 1) Then
                Temp = CartPrices(Pos)
                CartPrices(Pos) = CartPrices(Pos + 1)
                CartPrices(Pos + 1) = Temp
            End If
        Next Pos
    Next Pass
    
End Sub

Private Sub cmdTotal_Click()

For J = 1 To Items
        picResults.Print ShoppingCart(J); Tab(20); FormatCurrency(CartPrices(J))
        Total = Total + CartPrices(J)
    Next J
    
    picResults.Print Tab(20); FormatCurrency(Total)
End Sub

