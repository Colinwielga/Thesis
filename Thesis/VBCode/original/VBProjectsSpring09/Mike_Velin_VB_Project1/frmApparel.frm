VERSION 5.00
Begin VB.Form frmApparel 
   BackColor       =   &H000000C0&
   Caption         =   "frmApparel"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   Picture         =   "frmApparel.frx":0000
   ScaleHeight     =   10155
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHoodie 
      Caption         =   "Hoodie"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      TabIndex        =   39
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton cmdSkirt 
      Caption         =   "Skirt"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   38
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton cmdTop 
      Caption         =   "Top"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   37
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton cmdDress 
      Caption         =   "Dress"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   36
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton cmdPajamas 
      Caption         =   "Pajamas"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   35
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton cmdSwimsuit 
      Caption         =   "Swimsuit"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      TabIndex        =   34
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton cmdSocks 
      Caption         =   "Socks"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   33
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton cmdUnderwear 
      Caption         =   "Underwear"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   32
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton cmdShorts 
      Caption         =   "Shorts"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   31
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton cmdDressPants 
      Caption         =   "Dress Pants"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   30
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton cmdJeans 
      Caption         =   "Jeans"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      TabIndex        =   29
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdDressShirt 
      Caption         =   "Dress Shirt"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   28
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdPolo 
      Caption         =   "Polo"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   27
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdSweater 
      Caption         =   "Sweater"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   26
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdTShirt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "T-Shirt"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   25
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Price Check"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   24
      Top             =   6000
      Width           =   1455
   End
   Begin VB.TextBox txtPriceCheck 
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5760
      TabIndex        =   22
      Top             =   5880
      Width           =   3375
   End
   Begin VB.CommandButton cmdElectronics 
      Caption         =   "Electronics Department"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2400
      TabIndex        =   5
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home Department"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4680
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdShoes 
      Caption         =   "Shoe Department"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6960
      TabIndex        =   3
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdHomePage 
      Caption         =   "Home Page"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdToys 
      Caption         =   "Toy Department"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9240
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdShoppingCart 
      Caption         =   "Shopping Cart"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblEnter 
      BackColor       =   &H000000C0&
      Caption         =   "Enter a Number to Price Check the Item   -------->"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   23
      Top             =   5880
      Width           =   4095
   End
   Begin VB.Label lblHoodie 
      BackColor       =   &H000000C0&
      Caption         =   "15.  Hoodies"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   21
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label lblSkirt 
      BackColor       =   &H000000C0&
      Caption         =   "14.  Skirts"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   20
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label lblTop 
      BackColor       =   &H000000C0&
      Caption         =   "13.  Tops"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   19
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label lblDress 
      BackColor       =   &H000000C0&
      Caption         =   "12.  Dresses"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label lblPajamas 
      BackColor       =   &H000000C0&
      Caption         =   "11.  Pajamas"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label lblSwimsuit 
      BackColor       =   &H000000C0&
      Caption         =   "10.  Swimsuits"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   16
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblSocks 
      BackColor       =   &H000000C0&
      Caption         =   "9.  Socks"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   15
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label lblUnderwear 
      BackColor       =   &H000000C0&
      Caption         =   "8.  Underwear"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   14
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label lblShorts 
      BackColor       =   &H000000C0&
      Caption         =   "7.  Shorts"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lblDressPants 
      BackColor       =   &H000000C0&
      Caption         =   "6.  Dress Pants"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblJeans 
      BackColor       =   &H000000C0&
      Caption         =   "5.  Jeans"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   11
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblDressShirt 
      BackColor       =   &H000000C0&
      Caption         =   "4.  Dress Shirts"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   10
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label lblPolo 
      BackColor       =   &H000000C0&
      Caption         =   "3.  Polos"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblSweater 
      BackColor       =   &H000000C0&
      Caption         =   "2.  Sweaters"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label lblTShirt 
      BackColor       =   &H000000C0&
      Caption         =   "1.  T-Shirts"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label lblApparel 
      BackColor       =   &H000000C0&
      Caption         =   " Apparel Department"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1920
      TabIndex        =   6
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmApparel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCheck_Click()
        Dim CTR As Integer, Apparel(1 To 50) As String, ApparelNumber(1 To 50) As Integer, ApparelPrice(1 To 50) As Single, Found As Boolean, Pos As Integer, Number As String
        Open App.Path & "\Apparel.txt" For Input As #1
        CTR = 0
        Do While Not EOF(1)
            CTR = CTR + 1
            Input #1, ApparelNumber(CTR), Apparel(CTR), ApparelPrice(CTR)
        Loop
        Close #1
        Found = False
        Pos = 1
        Number = txtPriceCheck.Text
        Do While Not Found And Pos <= CTR
            If Number = ApparelNumber(Pos) Then
                MsgBox Apparel(Pos) & " cost " & ApparelPrice(Pos) & " dollars", , "Price Check"
                Found = True
            End If
            Pos = Pos + 1
        Loop
        If Found = False Then
            MsgBox "Try Again!", , "Error"
        End If
End Sub

Private Sub cmdDress_Click()
        Dresses = InputBox("Enter how many Dresses you would like to buy", "Dresses")
End Sub

Private Sub cmdDressPants_Click()
        DressPants = InputBox("Enter how many Dress Pants you would like to buy", "Dress Pants")
End Sub

Private Sub cmdDressShirt_Click()
        DressShirts = InputBox("Enter how many Dress Shirts you would like to buy", "Dress Shirts")
End Sub

Private Sub cmdElectronics_Click()
        frmApparel.Hide
        frmElectronics.Show
End Sub

Private Sub cmdHome_Click()
        frmApparel.Hide
        frmHome.Show
End Sub

Private Sub cmdHomePage_Click()
        frmApparel.Hide
        frmTarget.Show
End Sub

Private Sub cmdHoodie_Click()
        Hoodies = InputBox("Enter how many Hoodies you would like to buy", "Hoodies")
End Sub

Private Sub cmdJeans_Click()
        Jeans = InputBox("Enter how many Jeans you would like to buy", "Jeans")
End Sub

Private Sub cmdPajamas_Click()
        Pajamas = InputBox("Enter how many Pajamas you would like to buy", "Pajamas")
End Sub

Private Sub cmdPolo_Click()
        Polos = InputBox("Enter how many Polos you would like to buy", "Polos")
End Sub

Private Sub cmdShoes_Click()
        frmApparel.Hide
        frmShoes.Show
End Sub

Private Sub cmdShoppingCart_Click()
        frmApparel.Hide
        frmShoppingCart.Show
End Sub

Private Sub cmdShorts_Click()
        Shorts = InputBox("Enter how many Shorts you would like to buy", "Shorts")
End Sub

Private Sub cmdSkirt_Click()
        Skirts = InputBox("Enter how many Skirts you would like to buy", "Skirts")
End Sub

Private Sub cmdSocks_Click()
        Socks = InputBox("Enter how many Socks you would like to buy", "Socks")
End Sub

Private Sub cmdSweater_Click()
        Sweaters = InputBox("Enter how many Sweaters you would like to buy", "Sweaters")
End Sub

Private Sub cmdSwimsuit_Click()
        Swimsuits = InputBox("Enter how many Swimsuits you would like to buy", "Swimsuits")
End Sub

Private Sub cmdTop_Click()
        Tops = InputBox("Enter how many Tops you would like to buy", "Tops")
End Sub

Private Sub cmdToys_Click()
        frmApparel.Hide
        frmToys.Show
End Sub

Private Sub cmdTShirt_Click()
        TShirts = InputBox("Enter how many T-shirts you would like to buy", "T-Shirts")
End Sub

Private Sub cmdUnderwear_Click()
        Underwear = InputBox("Enter how much Underwear you would like to buy", "Underwear")
End Sub

Private Sub Form_Load()
        'Target, Corp
        'frmApparel.frm
        'Mike Velin
        'March 23rd, 2009
        'Providing Apparel products for the user to choose from
End Sub
