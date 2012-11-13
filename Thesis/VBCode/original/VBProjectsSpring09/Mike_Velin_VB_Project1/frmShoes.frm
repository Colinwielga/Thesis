VERSION 5.00
Begin VB.Form frmShoes 
   BackColor       =   &H000000C0&
   Caption         =   "frmShoes"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   Picture         =   "frmShoes.frx":0000
   ScaleHeight     =   9015
   ScaleWidth      =   11340
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
      Left            =   5520
      TabIndex        =   18
      Top             =   5880
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
      Left            =   9720
      TabIndex        =   17
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdCasual 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Casual Shoe"
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
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdDress 
      Caption         =   "Dress Shoe"
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
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdSandals 
      Caption         =   "Sandals"
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
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdAthletic 
      Caption         =   "Athletic Shoe"
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
      Left            =   7680
      TabIndex        =   13
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdFlats 
      Caption         =   "Flats"
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
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdPumps 
      Caption         =   "Pumps"
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
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdStappyHeels 
      Caption         =   "Strappy Heels"
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
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlatforms 
      Caption         =   "Platforms"
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
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdMoccasins 
      Caption         =   "Moccasins"
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
      Left            =   7680
      TabIndex        =   8
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdBoots 
      Caption         =   "Boots"
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
      Top             =   8280
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
      Left            =   6960
      TabIndex        =   4
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
   Begin VB.Label lblCasual 
      BackColor       =   &H000000C0&
      Caption         =   "1.  Casual Shoes"
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
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label lblDress 
      BackColor       =   &H000000C0&
      Caption         =   "2.  Dress Shoes"
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
      Left            =   2640
      TabIndex        =   28
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label lblSandals 
      BackColor       =   &H000000C0&
      Caption         =   "3.  Sandals"
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
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label lblAthletic 
      BackColor       =   &H000000C0&
      Caption         =   "4.  Athletic Shoes"
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
      Left            =   7680
      TabIndex        =   26
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label lblFlats 
      BackColor       =   &H000000C0&
      Caption         =   "5.  Flats"
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
      TabIndex        =   25
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label lblPumps 
      BackColor       =   &H000000C0&
      Caption         =   "6.  Pumps"
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
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lblStrappyHeels 
      BackColor       =   &H000000C0&
      Caption         =   "7.  Strappy Heels"
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
      Left            =   2640
      TabIndex        =   23
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label lblPlatforms 
      BackColor       =   &H000000C0&
      Caption         =   "8.  Platforms"
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
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label lblMoccasins 
      BackColor       =   &H000000C0&
      Caption         =   "9.  Moccasins"
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
      Left            =   7680
      TabIndex        =   21
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblBoots 
      BackColor       =   &H000000C0&
      Caption         =   "10.  Boots"
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
      TabIndex        =   20
      Top             =   4680
      Width           =   975
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
      Top             =   5880
      Width           =   4575
   End
   Begin VB.Label lblShoes 
      BackColor       =   &H000000C0&
      Caption         =   "Shoe Department"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmShoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApparel_Click()
        frmShoes.Hide
        frmApparel.Show
End Sub

Private Sub cmdAthletic_Click()
        AthleticShoes = InputBox("Enter how many Athletic Shoes you would like to buy", "Athletic Shoes")
End Sub

Private Sub cmdBoots_Click()
        Boots = InputBox("Enter how many Boots you would like to buy", "Boots")
End Sub

Private Sub cmdCasual_Click()
        CasualShoes = InputBox("Enter how many Casual Shoes you would like to buy", "Casual Shoes")
End Sub

Private Sub cmdCheck_Click()
        Dim CTR As Integer, Shoes(1 To 50) As String, ShoesNumber(1 To 50) As Integer, ShoesPrice(1 To 50) As Single, Found As Boolean, Pos As Integer, Number As String
        Open App.Path & "\Shoes.txt" For Input As #1
        CTR = 0
        Do While Not EOF(1)
            CTR = CTR + 1
            Input #1, ShoesNumber(CTR), Shoes(CTR), ShoesPrice(CTR)
        Loop
        Close #1
        Found = False
        Pos = 1
        Number = txtPriceCheck.Text
        Do While Not Found And Pos <= CTR
            If Number = ShoesNumber(Pos) Then
                MsgBox Shoes(Pos) & " cost " & ShoesPrice(Pos) & " dollars", , "Price Check"
                Found = True
            End If
            Pos = Pos + 1
        Loop
        If Found = False Then
            MsgBox "Try Again!", , "Error"
        End If
End Sub

Private Sub cmdDress_Click()
        DressShoes = InputBox("Enter how many Dress Shoes you would like to buy", "Dress Shoes")
End Sub

Private Sub cmdElectronics_Click()
        frmShoes.Hide
        frmElectronics.Show
End Sub

Private Sub cmdFlats_Click()
        Flats = InputBox("Enter how many Flats you would like to buy", "Flats")
End Sub

Private Sub cmdHome_Click()
        frmShoes.Hide
        frmHome.Show
End Sub

Private Sub cmdHomePage_Click()
        frmShoes.Hide
        frmTarget.Show
End Sub

Private Sub cmdMoccasins_Click()
        Moccasins = InputBox("Enter how many Moccasins you would like to buy", "Moccasins")
End Sub

Private Sub cmdPlatforms_Click()
        Platforms = InputBox("Enter how many Platforms you would like to buy", "Platforms")
End Sub

Private Sub cmdPumps_Click()
        Pumps = InputBox("Enter how many Pumps you would like to buy", "Pumps")
End Sub

Private Sub cmdSandals_Click()
        Sandals = InputBox("Enter how many Sandals you would like to buy", "Sandals")
End Sub

Private Sub cmdShoppingCart_Click()
        frmShoes.Hide
        frmShoppingCart.Show
End Sub

Private Sub cmdStappyHeels_Click()
        StrappyHeels = InputBox("Enter how many Strappy Heels you would like to buy", "Strappy Heels")
End Sub

Private Sub cmdToys_Click()
        frmShoes.Hide
        frmToys.Show
End Sub

Private Sub Form_Load()
        'Target, Corp.
        'frmShoes.frm
        'Mike Velin
        'March 23rd, 2009
        'To provide various shoe products to the user
End Sub
