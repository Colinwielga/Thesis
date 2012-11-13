VERSION 5.00
Begin VB.Form frmElectronics 
   BackColor       =   &H000000C0&
   Caption         =   "frmElectronics"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   Picture         =   "frmElectronics.frx":0000
   ScaleHeight     =   8805
   ScaleWidth      =   11400
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
      TabIndex        =   32
      Top             =   5520
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
      TabIndex        =   31
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdcamera 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Camera"
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
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdTelevision 
      Caption         =   "Television"
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
      Left            =   2040
      TabIndex        =   29
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdComputer 
      Caption         =   "Computer"
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
      Left            =   3960
      TabIndex        =   28
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCellPhone 
      Caption         =   "Cell Phone"
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
      Left            =   5880
      TabIndex        =   27
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdiPod 
      Caption         =   "iPod"
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
      Left            =   7920
      TabIndex        =   26
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdHomeTheater 
      Caption         =   "Home Theater "
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
      Left            =   9960
      TabIndex        =   25
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdXbox 
      Caption         =   "Xbox"
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
      TabIndex        =   24
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlaystation 
      Caption         =   "Playstation"
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
      Left            =   2040
      TabIndex        =   23
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdWii 
      Caption         =   "Wii"
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
      Left            =   3960
      TabIndex        =   22
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdGame 
      Caption         =   "Game"
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
      Left            =   5880
      TabIndex        =   21
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdMovie 
      Caption         =   "Movie"
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
      Left            =   7920
      TabIndex        =   20
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdCD 
      Caption         =   "CD"
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
      Left            =   9960
      TabIndex        =   19
      Top             =   8040
      Width           =   1215
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
      BackColor       =   &H8000000E&
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
      TabIndex        =   33
      Top             =   5520
      Width           =   4215
   End
   Begin VB.Label lblCD 
      BackColor       =   &H000000C0&
      Caption         =   "12.  CD's"
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
      Left            =   9360
      TabIndex        =   18
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label lblMovie 
      BackColor       =   &H000000C0&
      Caption         =   "11.  Movies"
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
      Left            =   7440
      TabIndex        =   17
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label lblGame 
      BackColor       =   &H000000C0&
      Caption         =   "10.  Games"
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
      Left            =   5760
      TabIndex        =   16
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label lblWii 
      BackColor       =   &H000000C0&
      Caption         =   "9.  Wii's"
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
      Left            =   3960
      TabIndex        =   15
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label lblPlaystation 
      BackColor       =   &H000000C0&
      Caption         =   "8.  Playstations"
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
      Left            =   2040
      TabIndex        =   14
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label lblXbox 
      BackColor       =   &H000000C0&
      Caption         =   "7.  Xbox's"
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
      TabIndex        =   13
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label lblHomeTheater 
      BackColor       =   &H000000C0&
      Caption         =   "6.  Home Theaters"
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
      Left            =   9360
      TabIndex        =   12
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label lbliPod 
      BackColor       =   &H000000C0&
      Caption         =   "5.  iPods"
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
      Left            =   7440
      TabIndex        =   11
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lblCellPhone 
      BackColor       =   &H000000C0&
      Caption         =   "4.  Cell Phones"
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
      Left            =   5760
      TabIndex        =   10
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblComputer 
      BackColor       =   &H000000C0&
      Caption         =   "3.  Computers"
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
      Left            =   3960
      TabIndex        =   9
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label llblTelevision 
      BackColor       =   &H000000C0&
      Caption         =   "2.  Televisions"
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
      Left            =   2040
      TabIndex        =   8
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblCamera 
      BackColor       =   &H000000C0&
      Caption         =   "1.  Cameras"
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
      TabIndex        =   7
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lblElectronics 
      BackColor       =   &H000000C0&
      Caption         =   "Electronics Department"
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
      Left            =   1920
      TabIndex        =   6
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmElectronics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApparel_Click()
        frmElectronics.Hide
        frmApparel.Show
End Sub

Private Sub cmdcamera_Click()
        Cameras = InputBox("Enter how many Cameras you would like to buy", "Cameras")
End Sub

Private Sub cmdCD_Click()
        CDs = InputBox("Enter how many CD's you would like to buy", "CD's")
End Sub

Private Sub cmdCellPhone_Click()
        CellPhones = InputBox("Enter how many Cell Phones you would like to buy", "Cell Phones")
End Sub

Private Sub cmdCheck_Click()
        Dim CTR As Integer, Electronics(1 To 50) As String, ElectronicsNumber(1 To 50) As Integer, ElectronicsPrice(1 To 50) As Single, Found As Boolean, Pos As Integer, Number As String
        Open App.Path & "\Electronics.txt" For Input As #1
        CTR = 0
        Do While Not EOF(1)
            CTR = CTR + 1
            Input #1, ElectronicsNumber(CTR), Electronics(CTR), ElectronicsPrice(CTR)
        Loop
        Close #1
        Found = False
        Pos = 1
        Number = txtPriceCheck.Text
        Do While Not Found And Pos <= CTR
            If Number = ElectronicsNumber(Pos) Then
                MsgBox Electronics(Pos) & " cost " & ElectronicsPrice(Pos) & " dollars", , "Price Check"
                Found = True
            End If
            Pos = Pos + 1
        Loop
        If Found = False Then
            MsgBox "Try Again!", , "Error"
        End If
End Sub

Private Sub cmdComputer_Click()
        Computers = InputBox("Enter how many Computers you would like to buy", "Computers")
End Sub

Private Sub cmdGame_Click()
        Games = InputBox("Enter how many Games you would like to buy", "Games")
End Sub

Private Sub cmdHome_Click()
        frmElectronics.Hide
        frmHome.Show
End Sub

Private Sub cmdHomePage_Click()
        frmElectronics.Hide
        frmTarget.Show
End Sub

Private Sub cmdHomeTheater_Click()
        HomeTheaters = InputBox("Enter how many Home Theaters you would like to buy", "Home Theaters")
End Sub

Private Sub cmdiPod_Click()
        iPods = InputBox("Enter how many iPods you would like to buy", "iPods")
End Sub

Private Sub cmdMovie_Click()
        Movies = InputBox("Enter how many Movies you would like to buy", "Movies")
End Sub

Private Sub cmdPlaystation_Click()
        Playstations = InputBox("Enter how many Playstations you would like to buy", "Playstations")
End Sub

Private Sub cmdShoes_Click()
        frmElectronics.Hide
        frmShoes.Show
End Sub

Private Sub cmdShoppingCart_Click()
        frmElectronics.Hide
        frmShoppingCart.Show
End Sub

Private Sub cmdTelevision_Click()
        Televisions = InputBox("Enter how many Televisions you would like to buy", "Televisions")
End Sub

Private Sub cmdToys_Click()
        frmElectronics.Hide
        frmToys.Show
End Sub

Private Sub cmdWii_Click()
        Wiis = InputBox("Enter how many Wii's you would like to buy", "Wii's")
End Sub

Private Sub cmdXbox_Click()
        Xboxs = InputBox("Enter how many Xbox's you would like to buy", "Xbox's")
End Sub

Private Sub Form_Load()
        'Target, Corp.
        'frmElectronics.frm
        'Mike Velin
        'March 23rd, 2009
        'To provide the user with possible electronics needs and the price of each tiem
End Sub
