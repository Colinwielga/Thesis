VERSION 5.00
Begin VB.Form frmToys 
   BackColor       =   &H000000C0&
   Caption         =   "frmToys"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   Picture         =   "frmToys.frx":0000
   ScaleHeight     =   8580
   ScaleWidth      =   11595
   StartUpPosition =   3  'Windows Default
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
      Left            =   5640
      TabIndex        =   18
      Top             =   5400
      Width           =   3375
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
      Left            =   9840
      TabIndex        =   17
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdCribToy 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Crib Toy"
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
      TabIndex        =   16
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdStrollerToy 
      Caption         =   "Stroller Toy"
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
      TabIndex        =   15
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSwing 
      Caption         =   "Electronic Swing"
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
      TabIndex        =   14
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdDoll 
      Caption         =   "Doll"
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
      TabIndex        =   13
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdKitchen 
      Caption         =   "Play Kitchen"
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
      TabIndex        =   12
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdRidingToy 
      Caption         =   "Riding Toy"
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
      TabIndex        =   11
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdRadioFlyer 
      Caption         =   "RadioFlyer"
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
      TabIndex        =   10
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdLego 
      Caption         =   "Legos"
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
      TabIndex        =   9
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdBoardGame 
      Caption         =   "Board Game"
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
      TabIndex        =   8
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdNerf 
      Caption         =   "Nerf Toys"
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
      TabIndex        =   7
      Top             =   7800
      Width           =   1215
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
      Top             =   2160
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
      Left            =   9240
      TabIndex        =   4
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton cmdShoes 
      Caption         =   "Shoes Department"
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
      Top             =   2160
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
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton cmdApparel 
      Caption         =   "Apparel Department"
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
      TabIndex        =   1
      Top             =   2160
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
   Begin VB.Label lblCribToy 
      BackColor       =   &H000000C0&
      Caption         =   "1.  Crib Toys"
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
      TabIndex        =   29
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label lblStrollerToy 
      BackColor       =   &H000000C0&
      Caption         =   "2.  Stroller Toys"
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
      TabIndex        =   28
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label lblSwing 
      BackColor       =   &H000000C0&
      Caption         =   "3.  Electronic Swings"
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
      TabIndex        =   27
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label lblDoll 
      BackColor       =   &H000000C0&
      Caption         =   "4.  Dolls"
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
      TabIndex        =   26
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label lblKitchen 
      BackColor       =   &H000000C0&
      Caption         =   "5.  Play Kitchens"
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
      Left            =   9840
      TabIndex        =   25
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label lblRidingToy 
      BackColor       =   &H000000C0&
      Caption         =   "6.  Riding Toys"
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
      TabIndex        =   24
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblRadioFlyer 
      BackColor       =   &H000000C0&
      Caption         =   "7.  Radio Flyers"
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
      TabIndex        =   23
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label lblLego 
      BackColor       =   &H000000C0&
      Caption         =   "8.  Legos"
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
      TabIndex        =   22
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label lblBoardGame 
      BackColor       =   &H000000C0&
      Caption         =   "9.  Board Games"
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
      TabIndex        =   21
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label lblNerf 
      BackColor       =   &H000000C0&
      Caption         =   "10.  Nerf Toys"
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
      Left            =   9840
      TabIndex        =   20
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label lblEnter 
      BackColor       =   &H000000C0&
      Caption         =   "Enter a Number to Price Check the Item  --------->"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   19
      Top             =   5400
      Width           =   4935
   End
   Begin VB.Label lblToys 
      BackColor       =   &H000000C0&
      Caption         =   "Toy Department"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmToys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApparel_Click()
        frmToys.Hide
        frmApparel.Show
End Sub

Private Sub cmdBoardGame_Click()
        BoardGames = InputBox("Enter how many BoardGames you would like to buy", "Board Games")
End Sub

Private Sub cmdCheck_Click()
        Dim CTR As Integer, Toys(1 To 50) As String, ToysNumber(1 To 50) As Integer, ToysPrice(1 To 50) As Single, Found As Boolean, Pos As Integer, Number As String
        Open App.Path & "\Toys.txt" For Input As #1
        CTR = 0
        Do While Not EOF(1)
            CTR = CTR + 1
            Input #1, ToysNumber(CTR), Toys(CTR), ToysPrice(CTR)
        Loop
        Close #1
        Found = False
        Pos = 1
        Number = txtPriceCheck.Text
        Do While Not Found And Pos <= CTR
            If Number = ToysNumber(Pos) Then
                MsgBox Toys(Pos) & " cost " & ToysPrice(Pos) & " dollars", , "Price Check"
                Found = True
            End If
            Pos = Pos + 1
        Loop
        If Found = False Then
            MsgBox "Try Again!", , "Error"
        End If
End Sub

Private Sub cmdCribToy_Click()
        CribToys = InputBox("Enter how many Crib Toys you would like to buy", "Crib Toys")
End Sub

Private Sub cmdDoll_Click()
        Dolls = InputBox("Enter how many Dolls you would like to buy", "Dolls")
End Sub

Private Sub cmdElectronics_Click()
        frmToys.Hide
        frmElectronics.Show
End Sub

Private Sub cmdHome_Click()
        frmToys.Hide
        frmHome.Show
End Sub

Private Sub cmdHomePage_Click()
        frmToys.Hide
        frmTarget.Show
End Sub

Private Sub cmdKitchen_Click()
        PlayKitchens = InputBox("Enter how many Play Kitchens you would like to buy", "Play Kitchens")
End Sub

Private Sub cmdLego_Click()
        Legos = InputBox("Enter how many Legos you would like to buy", "Legos")
End Sub

Private Sub cmdNerf_Click()
        NerfToys = InputBox("Enter how many Nerf Toys you would like to buy", "Nerf Toys")
End Sub

Private Sub cmdRadioFlyer_Click()
        RadioFlyers = InputBox("Enter how many Radio Flyers you would like to buy", "Radio Flyers")
End Sub

Private Sub cmdRidingToy_Click()
        RidingToys = InputBox("Enter how many Riding Toys you would like to buy", "Riding Toys")
End Sub

Private Sub cmdShoes_Click()
        frmToys.Hide
        frmShoes.Show
End Sub

Private Sub cmdShoppingCart_Click()
        frmToys.Hide
        frmShoppingCart.Show
End Sub

Private Sub cmdStrollerToy_Click()
        StrollerToys = InputBox("Enter how many Stroller Toys you would like to buy", "Stroller Toys")
End Sub

Private Sub cmdSwing_Click()
        ElectronicSwings = InputBox("Enter how many Electronic Swings you would like to buy", "Electronic Swings")
End Sub

Private Sub Form_Load()
        'Target, Corp.
        'frmToys.frm
        'Mike Velin
        'March 23rd, 2009
        'To allow the user to browse the toy variety and choose an amount of specific toys
End Sub

Private Sub txtPriceCheck_Change()

End Sub
