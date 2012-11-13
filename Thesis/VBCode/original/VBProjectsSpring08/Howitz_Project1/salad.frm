VERSION 5.00
Begin VB.Form frmsalad 
   BackColor       =   &H0000FF00&
   Caption         =   "Form1"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8415
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   1095
      Left            =   360
      TabIndex        =   7
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdgarden 
      Caption         =   "Garden Salad"
      Height          =   735
      Left            =   2040
      TabIndex        =   6
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmddin 
      Caption         =   "Dinner Salad"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdSide 
      Caption         =   "Side Salad"
      Height          =   735
      Left            =   2040
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdPasta 
      Caption         =   "Pasta Salad"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdboss 
      Caption         =   "Boss salad"
      Height          =   735
      Left            =   2040
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdchef 
      Caption         =   "Chef Salad"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000E&
      Height          =   8535
      Left            =   4200
      ScaleHeight     =   8475
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "frmsalad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Project Name: Sexton Cash Register
'Form Name:  frmsalad
'Louis Howitz
'March 31, 2008
'All of the salad items are displayed on the buttons.  The
'buttons will print the item and price.

Private Sub cmdBack_Click()
    frmsalad.Hide
    frmTill.Show
    
End Sub

Private Sub cmdboss_Click()
    
    Dim Boss(1 To 6) As String
    Dim Price(1 To 6) As Single
    Open App.Path & "\Salad.txt" For Input As #1
    Do Until EOF(1)
        Input #1, Boss(6), Price(6)
        picResults.Print Boss(6); Tab(20); FormatCurrency(Price(6))
        frmPay.picResults.Print Boss(6); Tab(20); FormatCurrency(Price(6))
    Loop
    Close #1
End Sub


Private Sub cmdchef_Click()

    
   Items = Items + 1
        ShoppingCart(Items) = Salad(1)
        CartPrices(Items) = SaladPrice(1)
        picResults.Print Salad(1); Tab(20); FormatCurrency(SaladPrice(1))
        
    Close #1
End Sub

Private Sub cmddin_Click()

    Items = Items + 1
        ShoppingCart(Items) = Salad(2)
        CartPrices(Items) = SaladPrice(2)
        picResults.Print Salad(2); Tab(20); FormatCurrency(SaladPrice(2))
        
    Close #1
End Sub

Private Sub cmdgarden_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Salad(4)
        CartPrices(Items) = SaladPrice(4)
        picResults.Print Salad(4); Tab(20); FormatCurrency(SaladPrice(4))
        
    Close #1
End Sub

Private Sub cmdPasta_Click()

    Items = Items + 1
        ShoppingCart(Items) = Salad(3)
        CartPrices(Items) = SaladPrice(3)
        picResults.Print Salad(3); Tab(20); FormatCurrency(SaladPrice(3))
        
    Close #1
End Sub

Private Sub cmdSide_Click()
    
    Items = Items + 1
        ShoppingCart(Items) = Salad(5)
        CartPrices(Items) = SaladPrice(5)
        picResults.Print Salad(5); Tab(20); FormatCurrency(SaladPrice(5))
        
    Close #1
End Sub

Private Sub Form_Load()
'The file is loaded from an array.

Dim CTR As Integer
CTR = 0
    
    Open App.Path & "\Salad.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Salad(CTR), SaladPrice(CTR)
    Loop
    Close #1
End Sub
