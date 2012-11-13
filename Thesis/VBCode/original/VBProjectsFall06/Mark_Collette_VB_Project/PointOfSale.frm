VERSION 5.00
Begin VB.Form PointOfSale 
   BackColor       =   &H80000001&
   Caption         =   "Ace Hardware (Point of Sale)"
   ClientHeight    =   8745
   ClientLeft      =   2340
   ClientTop       =   1215
   ClientWidth     =   10980
   LinkTopic       =   "Form2"
   ScaleHeight     =   8745
   ScaleWidth      =   10980
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000040C0&
      Caption         =   "Go Back To Status"
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   9600
      Width           =   1575
   End
   Begin VB.CommandButton cmdShipping 
      BackColor       =   &H0000C000&
      Caption         =   "Shipping And Handling"
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00808000&
      Caption         =   ".........."
      Height          =   495
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton cmdSpecialOrder 
      BackColor       =   &H0000C000&
      Caption         =   "Special Order"
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton cmdParts 
      BackColor       =   &H0000C000&
      Caption         =   "Parts And Materials"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton cmdLabor 
      BackColor       =   &H0000C000&
      Caption         =   "Labor"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton cmdComplete 
      BackColor       =   &H00C000C0&
      Caption         =   "Complete Transaction"
      Height          =   615
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   9720
      Width           =   1935
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H00C000C0&
      Caption         =   "Total"
      Height          =   615
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   8880
      Width           =   1935
   End
   Begin VB.CommandButton cmdSnowShovel 
      BackColor       =   &H0080FF80&
      Caption         =   "Snow Shovel  16 in. Aluminum"
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton cmdLawnRake 
      BackColor       =   &H0080FF80&
      Caption         =   "Lawn Rake"
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton cmdFireWood 
      BackColor       =   &H0080FF80&
      Caption         =   "Fire Wood  5 lb."
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrunner 
      BackColor       =   &H0080FF80&
      Caption         =   "Hand Prunner"
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton cmdWeedKiller 
      BackColor       =   &H0080FF80&
      Caption         =   "Weed Killer"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton cmdGrassSeed 
      BackColor       =   &H0080FF80&
      Caption         =   "Grass Seed"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton cmdFertilizer 
      BackColor       =   &H0080FF80&
      Caption         =   "Phos Free Lawn Fertilizer"
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdSprinkler 
      BackColor       =   &H0080FF80&
      Caption         =   "Sprinkler"
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdGardenHose 
      BackColor       =   &H0080FF80&
      Caption         =   "Garden Hose"
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdFirepit 
      BackColor       =   &H0080FF80&
      Caption         =   "Copper Firepit"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdRollers 
      BackColor       =   &H0080FF80&
      Caption         =   "Paint Rollers  3pk."
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmdBrush2in 
      BackColor       =   &H0080FF80&
      Caption         =   "2 In. Paint Brush"
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmdBrush1in 
      BackColor       =   &H0080FF80&
      Caption         =   "1 In. Paint Brush"
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmdMineralSpirits 
      BackColor       =   &H0080FF80&
      Caption         =   "Mineral Spirits  Qt."
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmdPaintThinner 
      BackColor       =   &H0080FF80&
      Caption         =   "Paint Thinner  Qt."
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmdPolyurethane 
      BackColor       =   &H0080FF80&
      Caption         =   "Oil Based Polyurethane"
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdWoodStain 
      BackColor       =   &H0080FF80&
      Caption         =   "Oil Based Wood Stain"
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdPaint5Gal 
      BackColor       =   &H0080FF80&
      Caption         =   "Latex Paint 5 Gal."
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdPaintGal 
      BackColor       =   &H0080FF80&
      Caption         =   "Latex Paint Gal."
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdPaintQt 
      BackColor       =   &H0080FF80&
      Caption         =   "Latex Paint Qt."
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdHammer 
      BackColor       =   &H0080FF80&
      Caption         =   "16 oz. Claw Hammer"
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdSandpaperSheets 
      BackColor       =   &H0080FF80&
      Caption         =   "Sandpaper Sheets  4pk."
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdPowerSander 
      BackColor       =   &H0080FF80&
      Caption         =   "Power Sander"
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdSocketWrenchSet 
      BackColor       =   &H0080FF80&
      Caption         =   "Socket and Wrench Set"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdScrewdriverSet 
      BackColor       =   &H0080FF80&
      Caption         =   "Screwdriver Set"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdSawBlades 
      BackColor       =   &H0080FF80&
      Caption         =   "Saw Blades  4pk."
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdReciprocatingSaw 
      BackColor       =   &H0080FF80&
      Caption         =   "Reciprocating Saw"
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdDrillBits 
      BackColor       =   &H0080FF80&
      Caption         =   "Drill Bits"
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdCordlessDrill 
      BackColor       =   &H0080FF80&
      Caption         =   "Cordless Drill"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdFlashlight 
      BackColor       =   &H0080FF80&
      Caption         =   "Flashlight"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdMachineScrews 
      BackColor       =   &H0080FF80&
      Caption         =   "Assorted Machine Screws  1 lb."
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdWoodScrews 
      BackColor       =   &H0080FF80&
      Caption         =   "Assorted Wood Screws  1 lb."
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdPictureHangers 
      BackColor       =   &H0080FF80&
      Caption         =   "Picture Hangers"
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdRope 
      BackColor       =   &H0080FF80&
      Caption         =   "Rope  50 ft."
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdShelvingUnit 
      BackColor       =   &H0080FF80&
      Caption         =   "Shelving Unit"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdWindowShade 
      BackColor       =   &H0080FF80&
      Caption         =   "Window Shade"
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdGarageDoorOpener 
      BackColor       =   &H0080FF80&
      Caption         =   "Garage Door Opener"
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdPadlock 
      BackColor       =   &H0080FF80&
      Caption         =   "Padlock"
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdDoorLock 
      BackColor       =   &H0080FF80&
      Caption         =   "Door Lock"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdSmokeDetector 
      BackColor       =   &H0080FF80&
      Caption         =   "Smoke Detector"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.PictureBox picItemsSold 
      BackColor       =   &H80000005&
      Height          =   7935
      Left            =   10080
      ScaleHeight     =   7875
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton cmdViewReports 
      BackColor       =   &H000040C0&
      Caption         =   "View Reports"
      Height          =   855
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9600
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000040C0&
      Caption         =   "Exit Application"
      Height          =   855
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9600
      Width           =   1575
   End
   Begin VB.PictureBox picTotal 
      BackColor       =   &H80000005&
      Height          =   1215
      Left            =   13440
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   8160
      Width           =   1575
   End
   Begin VB.PictureBox picPricesSold 
      BackColor       =   &H80000005&
      Height          =   7935
      Left            =   13440
      ScaleHeight     =   7875
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lbl 
      Caption         =   "  Subtotal:                                 Tax:                Total Due: "
      Height          =   1215
      Left            =   12600
      TabIndex        =   50
      Top             =   8160
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Miscellaneous:"
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label lblPaint 
      BackColor       =   &H0000FFFF&
      Caption         =   "Paint:"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label lblTools 
      BackColor       =   &H0000FFFF&
      Caption         =   "Tools:"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblHardware 
      BackColor       =   &H0000FFFF&
      Caption         =   "Hardware:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "PointOfSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LaborPrice As Single
Dim ItemPrice As Single
Private Sub cmdBack_Click()
        'this button jumps from the point of sale screen to the status screen
        PointOfSale.Visible = False
        Status.Visible = True
        
End Sub

Private Sub cmdBrush1in_Click()
        'input product price and calculate running subtotal
        ItemPrice = 6.49
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "1 In. Paint Brush"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdBrush2in_Click()
        'input product price and calculate running subtotal
        ItemPrice = 12.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "2 In. Paint Brush"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdClear_Click()
        'this button clears the picture boxes displaying items, prices, and total
        picItemsSold.Cls
        picPricesSold.Cls
        picTotal.Cls
        Subtotal = 0
End Sub

Private Sub cmdComplete_Click()
        'this button prompts for a form of payment and completes a transaction
        If Subtotal = 0 Then
            MsgBox "There Is No Transaction To Complete", , "Error"
            TransactionDetails.Visible = False
        Else
            PointOfSale.Visible = True
            TransactionDetails.Visible = True
            Status.Visible = False
        End If
                        
End Sub

Private Sub cmdCordlessDrill_Click()
        'input product price and calculate running subtotal
        ItemPrice = 129.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Cordless Drill"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdDoorLock_Click()
        'input product price and calculate running subtotal
        ItemPrice = 19.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Door Lock"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdDrillBits_Click()
        'input product price and calculate running subtotal
        ItemPrice = 7.49
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Drill Bits"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdFertilizer_Click()
        'input product price and calculate running subtotal
        ItemPrice = 9.49
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Phos Free Lawn Fertilizer"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdFirepit_Click()
        'input product price and calculate running subtotal
        ItemPrice = 69.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Copper Firepit"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdFireWood_Click()
        'input product price and calculate running subtotal
        ItemPrice = 4.59
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Fire Wood 5lb."
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdFlashlight_Click()
        'input product price and calculate running subtotal
        ItemPrice = 5.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Flashlight"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdGarageDoorOpener_Click()
        'input product price and calculate running subtotal
        ItemPrice = 13.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Garage Door Opener"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdGardenHose_Click()
        'input product price and calculate running subtotal
        ItemPrice = 16.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Garden Hose"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdGrassSeed_Click()
        'input product price and calculate running subtotal
        ItemPrice = 10.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Grass Seed"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdHammer_Click()
        'input product price and calculate running subtotal
        ItemPrice = 11.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "16 oz. Claw Hammer"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdLabor_Click()
        'input product price and calculate running subtotal
        LaborPrice = InputBox("Please Enter An Amount", "Input")
        Subtotal = Subtotal + LaborPrice
        picItemsSold.Print "Labor"
        picPricesSold.Print FormatCurrency(LaborPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdLawnRake_Click()
        'input product price and calculate running subtotal
        ItemPrice = 8.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Lawn Rake"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdMachineScrews_Click()
        'input product price and calculate running subtotal
        ItemPrice = 5.49
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Assorted Machine Screws"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdMineralSpirits_Click()
        'input product price and calculate running subtotal
        ItemPrice = 2.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Mineral Spirits Qt."
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdPadlock_Click()
        'input product price and calculate running subtotal
        ItemPrice = 6.49
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Padlock"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdPaint5Gal_Click()
        'input product price and calculate running subtotal
        ItemPrice = 74.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Latex Paint 5 Gal."
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdPaintGal_Click()
        'input product price and calculate running subtotal
        ItemPrice = 15.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Latex Paint Gal."
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdPaintQt_Click()
        'input product price and calculate running subtotal
        ItemPrice = 9.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Latex Paint Qt."
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdPaintThinner_Click()
        'input product price and calculate running subtotal
        ItemPrice = 3.49
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Paint Thinner Qt."
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdParts_Click()
        'input product price and calculate running subtotal
        ItemPrice = InputBox("Please Enter An Amount", "Input")
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Parts & Materials"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdPictureHangers_Click()
        'input product price and calculate running subtotal
        ItemPrice = 2.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Picture Hangers"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdPolyurethane_Click()
        'input product price and calculate running subtotal
        ItemPrice = 11.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Oil Based Polyurethane Qt."
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdPowerSander_Click()
        'input product price and calculate running subtotal
        ItemPrice = 54.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Power Sander"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdPrunner_Click()
        'input product price and calculate running subtotal
        ItemPrice = 13.49
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Hand Prunner"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdQuit_Click()
        End
End Sub

Private Sub cmdReciprocatingSaw_Click()
        'input product price and calculate running subtotal
        ItemPrice = 99.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Reciprocating Saw"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdRollers_Click()
        'input product price and calculate running subtotal
        ItemPrice = 7.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Paint Rollers 3pk."
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdRope_Click()
        'input product price and calculate running subtotal
        ItemPrice = 8.49
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Rope 50ft."
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdSandpaperSheets_Click()
        'input product price and calculate running subtotal
        ItemPrice = 7.49
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Sandpaper Sheets 4 pk."
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdSawBlades_Click()
        'input product price and calculate running subtotal
        ItemPrice = 9.49
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Saw Blades 4pk."
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdScrewdriverSet_Click()
        'input product price and calculate running subtotal
        ItemPrice = 21.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Screwdriver Set"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdShelvingUnit_Click()
        'input product price and calculate running subtotal
        ItemPrice = 19.49
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Shelving Unit"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdShipping_Click()
        'input product price and calculate running subtotal
        ItemPrice = InputBox("Please Enter An Amount", "Input")
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Shipping & Handling"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdSmokeDetector_Click()
        'input product price and calculate running subtotal
        ItemPrice = 14.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Smoke Detector"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdSnowShovel_Click()
        'input product price and calculate running subtotal
        ItemPrice = 16.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "16in. Aluminum Snow Shovel"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdSocketWrenchSet_Click()
        'input product price and calculate running subtotal
        ItemPrice = 32.49
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Socket & Wrench Set"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdSpecialOrder_Click()
        'input product price and calculate running subtotal
        ItemPrice = InputBox("Please Enter An Amount", "Input")
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Special Order"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdSprinkler_Click()
        'input product price and calculate running subtotal
        ItemPrice = 11.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Sprinkler"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Public Sub cmdTotal_Click()
        'Calculate Tax based on subtotal and add to get Total
        Tax = (Subtotal - LaborPrice) * 0.07
        Total = Subtotal + Tax
        If Subtotal = 0 Then
            picTotal.Print "    "
            picTotal.Print "No Items To"
            picTotal.Print "Total Up"
        Else
            picTotal.Print "    "
            picTotal.Print FormatCurrency(Tax, 2)
            picTotal.Print "    "
            picTotal.Print FormatCurrency(Total, 2)
        End If
End Sub

Private Sub cmdViewReports_Click()
        'jump from point of sale to inventory info
        PointOfSale.Visible = False
        Status.Visible = False
        InventoryInfo.Visible = True
        
End Sub

        
Private Sub cmdWeedKiller_Click()
        'input product price and calculate running subtotal
        ItemPrice = 8.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Weed Killer"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdWindowShade_Click()
        'input product price and calculate running subtotal
        ItemPrice = 12.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Window Shade"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdWoodScrews_Click()
        'input product price and calculate running subtotal
        ItemPrice = 4.99
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Assorted Wood Screws"
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub

Private Sub cmdWoodStain_Click()
        'input product price and calculate running subtotal
        ItemPrice = 12.49
        Subtotal = Subtotal + ItemPrice
        picItemsSold.Print "Oil Based Wood Stain Qt."
        picPricesSold.Print FormatCurrency(ItemPrice, 2)
        picTotal.Cls
        picTotal.Print FormatCurrency(Subtotal, 2)
End Sub



'RetailPOSandInventoryControl program; PointOfSale form
'this code was written on Tuesday, October 31, 2006
'this code was edited on Wednesday, November 1, 2006
'this code was revised on Thurs, November 2, 2006
'written by Mark Collette
'adapted from code from a similar project in lab 7
'the purpose of this form is to take input from the user in the form of command buttons
'the input was then used to calculate a running total and run a transaction as in a store
'each subroutine either reads in a price and calculates the running total, calculates a final total, jumps from one form to another, or clears data in picture boxes

