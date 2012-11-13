VERSION 5.00
Begin VB.Form frmShoppingCart 
   BackColor       =   &H000000C0&
   Caption         =   "frmShoppingCart"
   ClientHeight    =   14340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13665
   LinkTopic       =   "Form1"
   Picture         =   "frmShoppingCart.frx":0000
   ScaleHeight     =   14340
   ScaleWidth      =   13665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Shopping Cart"
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
      Left            =   3720
      TabIndex        =   11
      Top             =   12480
      Width           =   2775
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display Shopping Cart"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   12480
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
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
      Left            =   10800
      TabIndex        =   9
      Top             =   12480
      Width           =   2775
   End
   Begin VB.CommandButton cmdCheckOut 
      Caption         =   "Check Out"
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
      Left            =   7320
      TabIndex        =   8
      Top             =   12480
      Width           =   2775
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   8655
      Left            =   120
      ScaleHeight     =   8595
      ScaleWidth      =   13395
      TabIndex        =   7
      Top             =   3480
      Width           =   13455
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
      TabIndex        =   5
      Top             =   2040
      Width           =   2055
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
      TabIndex        =   4
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
      Left            =   9240
      TabIndex        =   3
      Top             =   2040
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdToys 
      Caption         =   "Toys Department"
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
      Left            =   11520
      TabIndex        =   0
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblShoppingCart 
      BackColor       =   &H000000C0&
      Caption         =   "Shopping Cart"
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
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmShoppingCart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
        Dim SubtotalApparel As Integer, SubtotalElectronics As Integer, SubtotalHome As Integer, SubtotalShoes As Integer, SubtotalToys As Integer, TotalProduct As Integer
        Dim SubtotalApparel2 As Single, SubtotalElectronics2 As Single, SubtotalHome2 As Single, SubtotalShoes2 As Single, SubtotalToys2 As Single, TotalPrice As Single

Private Sub cmdCheckOut_Click()
        Dim Shipping As Single, Tax As Single, TotalPrice2 As Single
        Shipping = 7.99
        Tax = TotalPrice * 0.07
        If TotalPrice > 50# Then
            Shipping = 0
        End If
        TotalPrice2 = TotalPrice + Tax + Shipping
        picResults.Print "Subtotal:"; Tab(80); FormatCurrency(TotalPrice)
        picResults.Print "Tax:"; Tab(80); FormatCurrency(Tax)
        picResults.Print "Shipping:"; Tab(80); FormatCurrency(Shipping)
        picResults.Print Tab(80); "*********"
        picResults.Print "Total Amount Due:"; Tab(80); FormatCurrency(TotalPrice2)
End Sub

Private Sub cmdClear_Click()
        picResults.Cls
        SubtotalApparel = 0
        SubtotalElectronics = 0
        SubtotalHome = 0
        SubtotalShoes = 0
        SubtotalToys = 0
        TotalProduct = 0
End Sub

Private Sub cmdDisplay_Click()
        Dim TShirts2 As Single, Sweaters2 As Single, Polos2 As Single, DressShirts2 As Single, Jeans2 As Single, DressPants2 As Single, Shorts2 As Single, Underwear2 As Single, Socks2 As Single, Swimsuits2 As Single, Pajamas2 As Single, Dresses2 As Single, Tops2 As Single, Skirts2 As Single, Hoodies2 As Single
        Dim Cameras2 As Single, Televisions2 As Single, Computers2 As Single, CellPhones2 As Single, iPods2 As Single, HomeTheaters2 As Single, Xboxs2 As Single, Playstations2 As Single, Wiis2 As Single, Games2 As Single, Movies2 As Single, CDs2 As Single
        Dim CasualShoes2 As Single, DressShoes2 As Single, Sandals2 As Single, AthleticShoes2 As Single, Flats2 As Single, Pumps2 As Single, StrappyHeels2 As Single, Platforms2 As Single, Moccasins2 As Single, Boots2 As Single
        Dim CribToys2 As Single, StrollerToys2 As Single, ElectronicSwings2 As Single, Dolls2 As Single, PlayKitchens2 As Single, RidingToys2 As Single, RadioFlyers2 As Single, Legos2 As Single, BoardGames2 As Single, NerfToys2 As Single
        Dim StorageBins2 As Single, LaundryBins2 As Single, Shelving2 As Single, Rugs2 As Single, Curtains2 As Single, TableLamps2 As Single, FloorLamps2 As Single, Pillows2 As Single, Mirrors2 As Single, Clocks2 As Single, Microwaves2 As Single, Blenders2 As Single, Refridgerators2 As Single, Toasters2 As Single, Dinnerware2 As Single, Drinkware2 As Single, Comforters2 As Single, Duvets2 As Single, Sheets2 As Single, Blankets2 As Single, Mattresses2 As Single, Towels2 As Single, Nightstands2 As Single, Dressers2 As Single, Desks2 As Single, OfficeChairs2 As Single, Bookcases2 As Single, Chairs2 As Single, Sofas2 As Single, DinningTables2 As Single, Grills2 As Single, PatioDinningSets2 As Single, Gazebos2 As Single, Umbrellas2 As Single, EntertainmentCenters2 As Single
        SubtotalApparel = 0
        SubtotalElectronics = 0
        SubtotalHome = 0
        SubtotalShoes = 0
        SubtotalToys = 0
        SubtotalApparel2 = 0
        SubtotalElectronics2 = 0
        SubtotalHome2 = 0
        SubtotalShoes2 = 0
        SubtotalToys2 = 0
        picResults.Cls
        picResults.Print "Apparel"
        picResults.Print "**********"
        If TShirts > 0 Then
            TShirts2 = TShirts * 12.99
            picResults.Print "T-Shirts:"; Tab(40); TShirts; Tab(80); FormatCurrency(TShirts2)
            SubtotalApparel = SubtotalApparel + TShirts
            SubtotalApparel2 = SubtotalApparel2 + TShirts2
        End If
        If Sweaters > 0 Then
            Sweaters2 = Sweaters * 23.99
            picResults.Print "Sweaters:"; Tab(40); Sweaters; Tab(80); FormatCurrency(Sweaters2)
            SubtotalApparel = SubtotalApparel + Sweaters
            SubtotalApparel2 = SubtotalApparel2 + Sweaters2
        End If
        If Polos > 0 Then
            Polos2 = Polos * 16.99
            picResults.Print "Polos:"; Tab(40); Polos; Tab(80); FormatCurrency(Polos2)
            SubtotalApparel = SubtotalApparel + Polos
            SubtotalApparel2 = SubtotalApparel2 + Polos2
        End If
        If DressShirts > 0 Then
            DressShirts2 = DressShirts * 26.99
            picResults.Print "Dress Shirts:"; Tab(40); DressShirts; Tab(80); FormatCurrency(DressShirts2)
            SubtotalApparel = SubtotalApparel + DressShirts
            SubtotalApparel2 = SubtotalApparel2 + DressShirts2
        End If
        If Jeans > 0 Then
            Jeans2 = Jeans * 34.99
            picResults.Print "Jeans:"; Tab(40); Jeans; Tab(80); FormatCurrency(Jeans2)
            SubtotalApparel = SubtotalApparel + Jeans
            SubtotalApparel2 = SubtotalApparel2 + Jeans2
        End If
        If DressPants > 0 Then
            DressPants2 = DressPants * 24.99
            picResults.Print "Dress Pants:"; Tab(40); DressPants; Tab(80); FormatCurrency(DressPants2)
            SubtotalApparel = SubtotalApparel + DressPants
            SubtotalApparel2 = SubtotalApparel2 + DressPants2
        End If
        If Shorts > 0 Then
            Shorts2 = Shorts * 10.99
            picResults.Print "Shorts:"; Tab(40); Shorts; Tab(80); FormatCurrency(Shorts2)
            SubtotalApparel = SubtotalApparel + Shorts
            SubtotalApparel2 = SubtotalApparel2 + Shorts2
        End If
        If Underwear > 0 Then
            Underwear2 = Underwear * 9.99
            picResults.Print "Underwear:"; Tab(40); Underwear; Tab(80); FormatCurrency(Underwear2)
            SubtotalApparel = SubtotalApparel + Underwear
            SubtotalApparel2 = SubtotalApparel2 + Underwear2
        End If
        If Socks > 0 Then
            Socks2 = Socks * 10.99
            picResults.Print "Socks:"; Tab(40); Socks; Tab(80); FormatCurrency(Socks2)
            SubtotalApparel = SubtotalApparel + Socks
            SubtotalApparel2 = SubtotalApparel2 + Socks2
        End If
        If Swimsuits > 0 Then
            Swimsuits2 = Swimsuits * 14.99
            picResults.Print "Swimsuits:"; Tab(40); Swimsuits; Tab(80); FormatCurrency(Swimsuits2)
            SubtotalApparel = SubtotalApparel + Swimsuits
            SubtotalApparel2 = SubtotalApparel2 + Swimsuits2
        End If
        If Pajamas > 0 Then
            Pajamas2 = Pajamas * 11.99
            picResults.Print "Pajamas:"; Tab(40); Pajamas; Tab(80); FormatCurrency(Pajamas2)
            SubtotalApparel = SubtotalApparel + Pajamas
            SubtotalApparel2 = SubtotalApparel2 + Pajamas2
        End If
        If Dresses > 0 Then
            Dresses2 = Dresses * 29.99
            picResults.Print "Dresses:"; Tab(40); Dresses; Tab(80); FormatCurrency(Dresses2)
            SubtotalApparel = SubtotalApparel + Dresses
            SubtotalApparel2 = SubtotalApparel2 + Dresses2
        End If
        If Tops > 0 Then
            Tops2 = Tops * 8#
            picResults.Print "Tops:"; Tab(40); Tops; Tab(80); FormatCurrency(Tops2)
            SubtotalApparel = SubtotalApparel + Tops
            SubtotalApparel2 = SubtotalApparel2 + Tops2
        End If
        If Skirts > 0 Then
            Skirts2 = Skirts * 21.99
            picResults.Print "Skirts:"; Tab(40); Skirts; Tab(80); FormatCurrency(Skirts2)
            SubtotalApparel = SubtotalApparel + Skirts
            SubtotalApparel2 = SubtotalApparel2 + Skirts2
        End If
        If Hoodies > 0 Then
            Hoodies2 = Hoodies * 15.99
            picResults.Print "Hoodies:"; Tab(40); Hoodies; Tab(80); FormatCurrency(Hoodies2)
            SubtotalApparel = SubtotalApparel + Hoodies
            SubtotalApparel2 = SubtotalApparel2 + Hoodies2
        End If
        picResults.Print "**********"
        picResults.Print "Amount of Apparel:"; Tab(40); SubtotalApparel; Tab(80); FormatCurrency(SubtotalApparel2)
        picResults.Print "**********"
        picResults.Print
        picResults.Print "Electronics"
        picResults.Print "**********"
        If Cameras > 0 Then
            Cameras2 = Cameras * 179.99
            picResults.Print "Cameras:"; Tab(40); Cameras; Tab(80); FormatCurrency(Cameras2)
            SubtotalElectronics = SubtotalElectronics + Cameras
            SubtotalElectronics2 = SubtotalElectronics2 + Cameras2
        End If
        If Televisions > 0 Then
            Televisions2 = Televisions * 499.99
            picResults.Print "Televisions:"; Tab(40); Televisions; Tab(80); FormatCurrency(Televisions2)
            SubtotalElectronics = SubtotalElectronics + Televisions
            SubtotalElectronics2 = SubtotalElectronics2 + Televisions2
        End If
        If Computers > 0 Then
            Computers2 = Computers * 1029.99
            picResults.Print "Computers:"; Tab(40); Computers; Tab(80); FormatCurrency(Computers2)
            SubtotalElectronics = SubtotalElectronics + Computers
            SubtotalElectronics2 = SubtotalElectronics2 + Computers2
        End If
        If CellPhones > 0 Then
            CellPhones2 = CellPhones * 109.99
            picResults.Print "CellPhones:"; Tab(40); CellPhones; Tab(80); FormatCurrency(CellPhones2)
            SubtotalElectronics = SubtotalElectronics + CellPhones
            SubtotalElectronics2 = SubtotalElectronics2 + CellPhones2
        End If
        If iPods > 0 Then
            iPods2 = iPods * 299.99
            picResults.Print "iPods:"; Tab(40); iPods; Tab(80); FormatCurrency(iPods2)
            SubtotalElectronics = SubtotalElectronics + iPods
            SubtotalElectronics2 = SubtotalElectronics2 + iPods2
        End If
        If HomeTheaters > 0 Then
            HomeTheaters2 = HomeTheaters * 199.99
            picResults.Print "HomeTheaters:"; Tab(40); HomeTheaters; Tab(80); FormatCurrency(HomeTheaters2)
            SubtotalElectronics = SubtotalElectronics + HomeTheaters
            SubtotalElectronics2 = SubtotalElectronics2 + HomeTheaters2
        End If
        If Xboxs > 0 Then
            Xboxs2 = Xboxs * 249.99
            picResults.Print "Xbox's:"; Tab(40); Xboxs; Tab(80); FormatCurrency(Xboxs2)
            SubtotalElectronics = SubtotalElectronics + Xboxs
            SubtotalElectronics2 = SubtotalElectronics2 + Xboxs2
        End If
        If Playstations > 0 Then
            Playstations2 = Playstations * 399.99
            picResults.Print "Playstations:"; Tab(40); Playstations; Tab(80); FormatCurrency(Playstations2)
            SubtotalElectronics = SubtotalElectronics + Playstations
            SubtotalElectronics2 = SubtotalElectronics2 + Playstations2
        End If
        If Wiis > 0 Then
            Wiis2 = Wiis * 339.99
            picResults.Print "Wii's:"; Tab(40); Wiis; Tab(80); FormatCurrency(Wiis2)
            SubtotalElectronics = SubtotalElectronics + Wiis
            SubtotalElectronics2 = SubtotalElectronics2 + Wiis2
        End If
        If Games > 0 Then
            Games2 = Games * 50#
            picResults.Print "Games:"; Tab(40); Games; Tab(80); FormatCurrency(Games2)
            SubtotalElectronics = SubtotalElectronics + Games
            SubtotalElectronics2 = SubtotalElectronics2 + Games2
        End If
        If Movies > 0 Then
            Movies2 = Movies * 19.99
            picResults.Print "Movies:"; Tab(40); Movies; Tab(80); FormatCurrency(Movies2)
            SubtotalElectronics = SubtotalElectronics + Movies
            SubtotalElectronics2 = SubtotalElectronics2 + Movies2
        End If
        If CDs > 0 Then
            CDs2 = CDs * 9.99
            picResults.Print "CD's:"; Tab(40); CDs; Tab(80); FormatCurrency(CDs2)
            SubtotalElectronics = SubtotalElectronics + CDs
            SubtotalElectronics2 = SubtotalElectronics2 + CDs2
        End If
        picResults.Print "**********"
        picResults.Print "Amount of Electronics:"; Tab(40); SubtotalElectronics; Tab(80); FormatCurrency(SubtotalElectronics2)
        picResults.Print "**********"
        picResults.Print
        picResults.Print "Home"
        picResults.Print "**********"
        If StorageBins > 0 Then
            StorageBins2 = StorageBins * 38.97
            picResults.Print "Storagte Bin's:"; Tab(40); StorageBins; Tab(80); FormatCurrency(StorageBins2)
            SubtotalHome = SubtotalHome + StorageBins
            SubtotalHome2 = SubtotalHome2 + StorageBins2
        End If
        If LaundryBins > 0 Then
            LaundryBins2 = LaundryBins * 49.99
            picResults.Print "Laundry Bin's:"; Tab(40); LaundryBins; Tab(80); FormatCurrency(LaundryBins2)
            SubtotalHome = SubtotalHome + LaundryBins
            SubtotalHome2 = SubtotalHome2 + LaundryBins2
        End If
        If Shelving > 0 Then
            Shelving2 = Shelving * 79.99
            picResults.Print "Shelving:"; Tab(40); Shelving; Tab(80); FormatCurrency(Shelving2)
            SubtotalHome = SubtotalHome + Shelving
            SubtotalHome2 = SubtotalHome2 + Shelving2
        End If
        If Rugs > 0 Then
            Rugs2 = Rugs * 149.99
            picResults.Print "Rugs:"; Tab(40); Rugs; Tab(80); FormatCurrency(Rugs2)
            SubtotalHome = SubtotalHome + Rugs
            SubtotalHome2 = SubtotalHome2 + Rugs2
        End If
        If Curtains > 0 Then
            Curtains2 = Curtains * 34.99
            picResults.Print "Curtains:"; Tab(40); Curtains; Tab(80); FormatCurrency(Curtains2)
            SubtotalHome = SubtotalHome + Curtains
            SubtotalHome2 = SubtotalHome2 + Curtains2
        End If
        If TableLamps > 0 Then
            TableLamps2 = TableLamps * 29.99
            picResults.Print "Table Lamps:"; Tab(40); TableLamps; Tab(80); FormatCurrency(TableLamps2)
            SubtotalHome = SubtotalHome + TableLamps
            SubtotalHome2 = SubtotalHome2 + TableLamps2
        End If
        If FloorLamps > 0 Then
            FloorLamps2 = FloorLamps * 54.99
            picResults.Print "Floor Lamps:"; Tab(40); FloorLamps; Tab(80); FormatCurrency(FloorLamps2)
            SubtotalHome = SubtotalHome + FloorLamps
            SubtotalHome2 = SubtotalHome2 + FloorLamps2
        End If
        If Pillows > 0 Then
            Pillows2 = Pillows * 12.99
            picResults.Print "Pillows:"; Tab(40); Pillows; Tab(80); FormatCurrency(Pillows2)
            SubtotalHome = SubtotalHome + Pillows
            SubtotalHome2 = SubtotalHome2 + Pillows2
        End If
        If Mirrors > 0 Then
            Mirrors2 = Mirrors * 99.99
            picResults.Print "Mirrors:"; Tab(40); Mirrors; Tab(80); FormatCurrency(Mirrors2)
            SubtotalHome = SubtotalHome + Mirrors
            SubtotalHome2 = SubtotalHome2 + Mirrors2
        End If
        If Clocks > 0 Then
            Clocks2 = Clocks * 64.99
            picResults.Print "Clocks:"; Tab(40); Clocks; Tab(80); FormatCurrency(Clocks2)
            SubtotalHome = SubtotalHome + Clocks
            SubtotalHome2 = SubtotalHome2 + Clocks2
        End If
        If Microwaves > 0 Then
            Microwaves2 = Microwaves * 109.99
            picResults.Print "Microwaves:"; Tab(40); Microwaves; Tab(80); FormatCurrency(Microwaves2)
            SubtotalHome = SubtotalHome + Microwaves
            SubtotalHome2 = SubtotalHome2 + Microwaves2
        End If
        If Blenders > 0 Then
            Blenders2 = Blenders * 44.99
            picResults.Print "Blenders:"; Tab(40); Blenders; Tab(80); FormatCurrency(Blenders2)
            SubtotalHome = SubtotalHome + Blenders
            SubtotalHome2 = SubtotalHome2 + Blenders2
        End If
        If Refridgerators > 0 Then
            Refridgerators2 = Refridgerators * 219.99
            picResults.Print "Refridgerators:"; Tab(40); Refridgerators; Tab(80); FormatCurrency(Refridgerators2)
            SubtotalHome = SubtotalHome + Refridgerators
            SubtotalHome2 = SubtotalHome2 + Refridgerators2
        End If
        If Toasters > 0 Then
            Toasters2 = Toasters * 39.99
            picResults.Print "Toasters:"; Tab(40); Toasters; Tab(80); FormatCurrency(Toasters2)
            SubtotalHome = SubtotalHome + Toasters
            SubtotalHome2 = SubtotalHome2 + Toasters2
        End If
        If Dinnerware > 0 Then
            Dinnerware2 = Dinnerware * 69.99
            picResults.Print "Dinnerware:"; Tab(40); Dinnerware; Tab(80); FormatCurrency(Dinnerware2)
            SubtotalHome = SubtotalHome + Dinnerware
            SubtotalHome2 = SubtotalHome2 + Dinnerware2
        End If
        If Drinkware > 0 Then
            Drinkware2 = Drinkware * 24.99
            picResults.Print "Drinkware:"; Tab(40); Drinkware; Tab(80); FormatCurrency(Drinkware2)
            SubtotalHome = SubtotalHome + Drinkware
            SubtotalHome2 = SubtotalHome2 + Drinkware2
        End If
        If Comforters > 0 Then
            Comforters2 = Comforters * 159.99
            picResults.Print "Comforters:"; Tab(40); Comforters; Tab(80); FormatCurrency(Comforters2)
            SubtotalHome = SubtotalHome + Comforters
            SubtotalHome2 = SubtotalHome2 + Comforters2
        End If
        If Duvets > 0 Then
            Duvets2 = Duvets * 89.99
            picResults.Print "Duvets:"; Tab(40); Duvets; Tab(80); FormatCurrency(Duvets2)
            SubtotalHome = SubtotalHome + Duvets
            SubtotalHome2 = SubtotalHome2 + Duvets2
        End If
        If Sheets > 0 Then
            Sheets2 = Sheets * 59.99
            picResults.Print "Sheets:"; Tab(40); Sheets; Tab(80); FormatCurrency(Sheets2)
            SubtotalHome = SubtotalHome + Sheets
            SubtotalHome2 = SubtotalHome2 + Sheets2
        End If
        If Blankets > 0 Then
            Blankets2 = Blankets * 37.99
            picResults.Print "Blankets:"; Tab(40); Blankets; Tab(80); FormatCurrency(Blankets2)
            SubtotalHome = SubtotalHome + Blankets
            SubtotalHome2 = SubtotalHome2 + Blankets2
        End If
        If Mattresses > 0 Then
            Mattresses2 = Mattresses * 1039.99
            picResults.Print "Mattresses:"; Tab(40); Mattresses; Tab(80); FormatCurrency(Mattresses2)
            SubtotalHome = SubtotalHome + Mattresses
            SubtotalHome2 = SubtotalHome2 + Mattresses2
        End If
        If Towels > 0 Then
            Towels2 = Towels * 17.49
            picResults.Print "Towels:"; Tab(40); Towels; Tab(80); FormatCurrency(Towels2)
            SubtotalHome = SubtotalHome + Towels
            SubtotalHome2 = SubtotalHome2 + Towels2
        End If
        If Nightstands > 0 Then
            Nightstands2 = Nightstands * 119.99
            picResults.Print "Nightstands:"; Tab(40); Nightstands; Tab(80); FormatCurrency(Nightstands2)
            SubtotalHome = SubtotalHome + Nightstands
            SubtotalHome2 = SubtotalHome2 + Nightstands2
        End If
        If Dressers > 0 Then
            Dressers2 = Dressers * 249.99
            picResults.Print "Dressers:"; Tab(40); Dressers; Tab(80); FormatCurrency(Dressers2)
            SubtotalHome = SubtotalHome + Dressers
            SubtotalHome2 = SubtotalHome2 + Dressers2
        End If
        If Desks > 0 Then
            Desks2 = Desks * 399.99
            picResults.Print "Desks:"; Tab(40); Desks; Tab(80); FormatCurrency(Desks2)
            SubtotalHome = SubtotalHome + Desks
            SubtotalHome2 = SubtotalHome2 + Desks2
        End If
        If OfficeChairs > 0 Then
            OfficeChairs2 = OfficeChairs * 139.99
            picResults.Print "Office Chairs:"; Tab(40); OfficeChairs; Tab(80); FormatCurrency(OfficeChairs2)
            SubtotalHome = SubtotalHome + OfficeChairs
            SubtotalHome2 = SubtotalHome2 + OfficeChairs2
        End If
        If Bookcases > 0 Then
            Bookcases2 = Bookcases * 279.99
            picResults.Print "Bookcases:"; Tab(40); Bookcases; Tab(80); FormatCurrency(Bookcases2)
            SubtotalHome = SubtotalHome + Bookcases
            SubtotalHome2 = SubtotalHome2 + Bookcases2
        End If
        If Chairs > 0 Then
            Chairs2 = Chairs * 329.99
            picResults.Print "Chairs:"; Tab(40); Chairs; Tab(80); FormatCurrency(Chairs2)
            SubtotalHome = SubtotalHome + Chairs
            SubtotalHome2 = SubtotalHome2 + Chairs2
        End If
        If Sofas > 0 Then
            Sofas2 = Sofas * 579.99
            picResults.Print "Sofas:"; Tab(40); Sofas; Tab(80); FormatCurrency(Sofas2)
            SubtotalHome = SubtotalHome + Sofas
            SubtotalHome2 = SubtotalHome2 + Sofas2
        End If
        If DinningTables > 0 Then
            DinningTables2 = DinningTables * 369.99
            picResults.Print "Dinning Tables:"; Tab(40); DinningTables; Tab(80); FormatCurrency(DinningTables2)
            SubtotalHome = SubtotalHome + DinningTables
            SubtotalHome2 = SubtotalHome2 + DinningTables2
        End If
        If Grills > 0 Then
            Grills2 = Grills * 1249.99
            picResults.Print "Grills:"; Tab(40); Grills; Tab(80); FormatCurrency(Grills2)
            SubtotalHome = SubtotalHome + Grills
            SubtotalHome2 = SubtotalHome2 + Grills2
        End If
        If PatioDinningSets > 0 Then
            PatioDinningSets2 = PatioDinningSets * 314.49
            picResults.Print "Patio Dinning Sets:"; Tab(40); PatioDinningSets; Tab(80); FormatCurrency(PatioDinningSets2)
            SubtotalHome = SubtotalHome + PatioDinningSets
            SubtotalHome2 = SubtotalHome2 + PatioDinningSets2
        End If
        If Gazebos > 0 Then
            Gazebos2 = Gazebos * 229.99
            picResults.Print "Gazebo's:"; Tab(40); Gazebos; Tab(80); FormatCurrency(Gazebos2)
            SubtotalHome = SubtotalHome + Gazebos
            SubtotalHome2 = SubtotalHome2 + Gazebos2
        End If
        If Umbrellas > 0 Then
            Umbrellas2 = Umbrellas * 439.99
            picResults.Print "Umbrellas:"; Tab(40); Umbrellas; Tab(80); FormatCurrency(Umbrellas2)
            SubtotalHome = SubtotalHome + Umbrellas
            SubtotalHome2 = SubtotalHome2 + Umbrellas2
        End If
        If EntertainmentCenters > 0 Then
            EntertainmentCenters2 = EntertainmentCenters * 249.99
            picResults.Print "Entertainment Centers:"; Tab(40); EntertainmentCenters; Tab(80); FormatCurrency(EntertainmentCenters2)
            SubtotalHome = SubtotalHome + EntertainmentCenters
            SubtotalHome2 = SubtotalHome2 + EntertainmentCenters2
        End If
        picResults.Print "**********"
        picResults.Print "Amount of Home Products:"; Tab(40); SubtotalHome; Tab(80); FormatCurrency(SubtotalHome2)
        picResults.Print "**********"
        picResults.Print
        picResults.Print "Shoes"
        picResults.Print "**********"
        If CasualShoes > 0 Then
            CasualShoes2 = CasualShoes * 39.99
            picResults.Print "Casual Shoes:"; Tab(40); CasualShoes; Tab(80); FormatCurrency(CasualShoes2)
            SubtotalShoes = SubtotalShoes + CasualShoes
            SubtotalShoes2 = SubtotalShoes2 + CasualShoes2
        End If
        If DressShoes > 0 Then
            DressShoes2 = DressShoes * 34.99
            picResults.Print "Dress Shoes:"; Tab(40); DressShoes; Tab(80); FormatCurrency(DressShoes2)
            SubtotalShoes = SubtotalShoes + DressShoes
            SubtotalShoes2 = SubtotalShoes2 + DressShoes2
        End If
        If Sandals > 0 Then
            Sandals2 = Sandals * 14.99
            picResults.Print "Sandals:"; Tab(40); Sandals; Tab(80); FormatCurrency(Sandals2)
            SubtotalShoes = SubtotalShoes + Sandals
            SubtotalShoes2 = SubtotalShoes2 + Sandals2
        End If
        If AthleticShoes > 0 Then
            AthleticShoes2 = AthleticShoes * 29.99
            picResults.Print "Athletic Shoes:"; Tab(40); AthleticShoes; Tab(80); FormatCurrency(AthleticShoes2)
            SubtotalShoes = SubtotalShoes + AthleticShoes
            SubtotalShoes2 = SubtotalShoes2 + AthleticShoes2
        End If
        If Flats > 0 Then
            Flats2 = Flats * 16.99
            picResults.Print "Flats:"; Tab(40); Flats; Tab(80); FormatCurrency(Flats2)
            SubtotalShoes = SubtotalShoes + Flats
            SubtotalShoes2 = SubtotalShoes2 + Flats2
        End If
        If Pumps > 0 Then
            Pumps2 = Pumps * 26.99
            picResults.Print "Pumps:"; Tab(40); Pumps; Tab(80); FormatCurrency(Pumps2)
            SubtotalShoes = SubtotalShoes + Pumps
            SubtotalShoes2 = SubtotalShoes2 + Pumps2
        End If
        If StrappyHeels > 0 Then
            StrappyHeels2 = StrappyHeels * 19.99
            picResults.Print "Strappy Heels:"; Tab(40); StrappyHeels; Tab(80); FormatCurrency(StrappyHeels2)
            SubtotalShoes = SubtotalShoes + StrappyHeels
            SubtotalShoes2 = SubtotalShoes2 + StrappyHeels2
        End If
        If Platforms > 0 Then
            Platforms2 = Platforms * 22.99
            picResults.Print "Platforms:"; Tab(40); Platforms; Tab(80); FormatCurrency(Platforms2)
            SubtotalShoes = SubtotalShoes + Platforms
            SubtotalShoes2 = SubtotalShoes2 + Platforms2
        End If
        If Moccasins > 0 Then
            Moccasins2 = Moccasins * 24.99
            picResults.Print "Moccasins:"; Tab(40); Moccasins; Tab(80); FormatCurrency(Moccasins2)
            SubtotalShoes = SubtotalShoes + Moccasins
            SubtotalShoes2 = SubtotalShoes2 + Moccasins2
        End If
        If Boots > 0 Then
            Boots2 = Boots * 59.99
            picResults.Print "Boots:"; Tab(40); Boots; Tab(80); FormatCurrency(Boots2)
            SubtotalShoes = SubtotalShoes + Boots
            SubtotalShoes2 = SubtotalShoes2 + Boots2
        End If
        picResults.Print "**********"
        picResults.Print "Amount of Shoes:"; Tab(40); SubtotalShoes; Tab(80); FormatCurrency(SubtotalShoes2)
        picResults.Print "**********"
        picResults.Print
        picResults.Print "Toys"
        picResults.Print "**********"
        If CribToys > 0 Then
            CribToys2 = CribToys * 19.99
            picResults.Print "Crib Toys:"; Tab(40); CribToys; Tab(80); FormatCurrency(CribToys2)
            SubtotalToys = SubtotalToys + CribToys
            SubtotalToys2 = SubtotalToys2 + CribToys2
        End If
        If StrollerToys > 0 Then
            StrollerToys2 = StrollerToys * 15.99
            picResults.Print "Stroller Toys:"; Tab(40); StrollerToys; Tab(80); FormatCurrency(StrollerToys2)
            SubtotalToys = SubtotalToys + StrollerToys
            SubtotalToys2 = SubtotalToys2 + StrollerToys2
        End If
        If ElectronicSwings > 0 Then
            ElectronicSwings2 = ElectronicSwings * 129.99
            picResults.Print "Electronic Swings:"; Tab(40); ElectronicSwings; Tab(80); FormatCurrency(ElectronicSwings2)
            SubtotalToys = SubtotalToys + ElectronicSwings
            SubtotalToys2 = SubtotalToys2 + ElectronicSwings2
        End If
        If Dolls > 0 Then
            Dolls2 = Dolls * 14.99
            picResults.Print "Dolls:"; Tab(40); Dolls; Tab(80); FormatCurrency(Dolls2)
            SubtotalToys = SubtotalToys + Dolls
            SubtotalToys2 = SubtotalToys2 + Dolls2
        End If
        If PlayKitchens > 0 Then
            PlayKitchens2 = PlayKitchens * 189#
            picResults.Print "Play Kitchens:"; Tab(40); PlayKitchens; Tab(80); FormatCurrency(PlayKitchens2)
            SubtotalToys = SubtotalToys + PlayKitchens
            SubtotalToys2 = SubtotalToys2 + PlayKitchens2
        End If
        If RidingToys > 0 Then
            RidingToys2 = RidingToys * 29.99
            picResults.Print "Riding Toys:"; Tab(40); RidingToys; Tab(80); FormatCurrency(RidingToys2)
            SubtotalToys = SubtotalToys + RidingToys
            SubtotalToys2 = SubtotalToys2 + RidingToys2
        End If
        If RadioFlyers > 0 Then
            RadioFlyers2 = RadioFlyers * 149.99
            picResults.Print "RadioFlyers:"; Tab(40); RadioFlyers; Tab(80); FormatCurrency(RadioFlyers2)
            SubtotalToys = SubtotalToys + RadioFlyers
            SubtotalToys2 = SubtotalToys2 + RadioFlyers2
        End If
        If Legos > 0 Then
            Legos2 = Legos * 24.99
            picResults.Print "Legos:"; Tab(40); Legos; Tab(80); FormatCurrency(Legos2)
            SubtotalToys = SubtotalToys + Legos
            SubtotalToys2 = SubtotalToys2 + Legos2
        End If
        If BoardGames > 0 Then
            BoardGames2 = BoardGames * 32.99
            picResults.Print "Board Games:"; Tab(40); BoardGames; Tab(80); FormatCurrency(BoardGames2)
            SubtotalToys = SubtotalToys + BoardGames
            SubtotalToys2 = SubtotalToys2 + BoardGames2
        End If
        If NerfToys > 0 Then
            NerfToys2 = NerfToys * 26.49
            picResults.Print "Nerf Toys:"; Tab(40); NerfToys; Tab(80); FormatCurrency(NerfToys2)
            SubtotalToys = SubtotalToys + NerfToys
            SubtotalToys2 = SubtotalToys2 + NerfToys2
        End If
        picResults.Print "**********"
        picResults.Print "Amount of Toys:"; Tab(40); SubtotalToys; Tab(80); FormatCurrency(SubtotalToys2)
        picResults.Print "**********"
        TotalProduct = SubtotalApparel + SubtotalElectronics + SubtotalHome + SubtotalShoes + SubtotalToys
        TotalPrice = SubtotalApparel2 + SubtotalElectronics2 + SubtotalHome2 + SubtotalShoes2 + SubtotalToys2
        picResults.Print "Total amount of products:"; Tab(40); TotalProduct; Tab(80); FormatCurrency(TotalPrice)
        picResults.Print
        picResults.Print
        picResults.Print
        picResults.Print
End Sub

Private Sub cmdElectronics_Click()
        frmShoppingCart.Hide
        frmElectronics.Show
End Sub

Private Sub cmdHome_Click()
        frmShoppingCart.Hide
        frmHome.Show
End Sub

Private Sub cmdHomePage_Click()
        frmShoppingCart.Hide
        frmTarget.Show
End Sub

Private Sub cmdQuit_Click()
        End
End Sub

Private Sub cmdShoes_Click()
        frmShoppingCart.Hide
        frmShoes.Show
End Sub

Private Sub cmdToys_Click()
        frmShoppingCart.Hide
        frmToys.Show
End Sub

Private Sub Form_Load()
        'Target, Corp.
        'frmShoppingCart.frm
        'Mike Velin
        'March 23rd, 2009
        'Allow the user to see items selected, along with check out
End Sub
