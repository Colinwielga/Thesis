VERSION 5.00
Begin VB.Form frmAccessories 
   BackColor       =   &H00000000&
   Caption         =   "Buy Accessories"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Home Page"
      Height          =   2655
      Left            =   8040
      TabIndex        =   11
      Top             =   5640
      Width           =   2295
   End
   Begin VB.CommandButton cmdCart 
      Caption         =   "Send selected Accessories to MY CART"
      Height          =   2655
      Left            =   5760
      TabIndex        =   10
      Top             =   5640
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Total"
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   8280
      Width           =   5535
   End
   Begin VB.CommandButton cmdTotalA 
      Caption         =   "Total for Accessories"
      Height          =   735
      Left            =   6000
      TabIndex        =   8
      Top             =   5640
      Width           =   1095
   End
   Begin VB.PictureBox picResultsaccessories 
      BackColor       =   &H000000FF&
      Height          =   2775
      Left            =   240
      ScaleHeight     =   2715
      ScaleWidth      =   5475
      TabIndex        =   5
      Top             =   5520
      Width           =   5535
   End
   Begin VB.Label Label9 
      BackColor       =   &H000000FF&
      Caption         =   "By: Ben Harper"
      Height          =   375
      Left            =   9120
      TabIndex        =   13
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H000000FF&
      Caption         =   "Cobra Double-Bass Pedals $129.99"
      Height          =   975
      Left            =   6240
      TabIndex        =   12
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000FF&
      Caption         =   "Drum Thrones $69.50"
      Height          =   855
      Left            =   9120
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      Caption         =   "Percussion Units $299.00"
      Height          =   735
      Left            =   9000
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.Image Thrones 
      Height          =   1800
      Left            =   7800
      Picture         =   "frmAccessories.frx":0000
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Image Pedals 
      Height          =   1590
      Left            =   4320
      Picture         =   "frmAccessories.frx":7DE2
      Top             =   3720
      Width           =   1800
   End
   Begin VB.Image Units 
      Height          =   1800
      Left            =   7800
      Picture         =   "frmAccessories.frx":11334
      Top             =   360
      Width           =   1035
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000FF&
      Caption         =   "All Drum Heads $45"
      Height          =   735
      Left            =   6360
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   "Remo Drum Sheild $250"
      Height          =   735
      Left            =   6360
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "Cox Travel Bags $99"
      Height          =   735
      Left            =   2160
      TabIndex        =   2
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Firth Drum Sticks           20 for $30"
      Height          =   855
      Left            =   2160
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Remo Drum Rings              $19.99"
      Height          =   855
      Left            =   2160
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
   Begin VB.Image Sticks 
      Height          =   975
      Left            =   120
      Picture         =   "frmAccessories.frx":174F6
      Top             =   2280
      Width           =   1800
   End
   Begin VB.Image Heads 
      Height          =   1755
      Left            =   4320
      Picture         =   "frmAccessories.frx":1D0A0
      Top             =   1680
      Width           =   1800
   End
   Begin VB.Image Sheild 
      Height          =   1245
      Left            =   4320
      Picture         =   "frmAccessories.frx":2756A
      Top             =   240
      Width           =   1800
   End
   Begin VB.Image Rings 
      Height          =   1710
      Left            =   120
      Picture         =   "frmAccessories.frx":2EA64
      Top             =   360
      Width           =   1800
   End
   Begin VB.Image Bags 
      Height          =   1470
      Left            =   120
      Picture         =   "frmAccessories.frx":38AF6
      Top             =   3600
      Width           =   1800
   End
End
Attribute VB_Name = "frmAccessories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Buy Drums Online (OnlineDrums.vbp)
'frmAccessories (frmAccessories)
'Ben Harper
'3/23/06
'This form allows the user to shop for accessories via pictures and general information given, such
'as the brand and price of each item. This form also allows the user to order drum sticks in sets of twenty for a discount price.
'this form will display only the accessory total before asking the user if it would like the items placed into his/her cart.







Private Sub cmdCart_Click()
    frmCart.Visible = True                 'sends accessory total to cart
    frmAccessories.Visible = False
    frmCart.picResults.Print "Accessories Purchases", "Total"
    frmCart.picResults.Print "*************************************************"
    frmCart.picResults.Print "You Accesories Total is: ", FormatCurrency(Accsum)
End Sub

Private Sub cmdClear_Click() 'clears accessory total
picResultsaccessories.Cls
Accsum = 0
End Sub

Private Sub cmdReturn_Click() 'returns to HomePage
frmAccessories.Visible = False
frmHomePage.Visible = True
End Sub

Private Sub cmdTotalA_Click()  'gets total for accessories and prints in accessories form
picResultsaccessories.Print "******************************************************"
picResultsaccessories.Print "Total Accessories: ", FormatCurrency(Accsum, 2)
cmdCart.Visible = True
End Sub

Private Sub Bags_Click()    'buys accessory and adds cost to total
picResultsaccessories.Print "Cox Gig Bags   ", FormatCurrency(99, 2)
Accsum = Accsum + 99
End Sub

Private Sub Rings_Click() 'buys accessory and adds cost to total
picResultsaccessories.Print "Remo Drum Rings", FormatCurrency(19.99, 2)
Accsum = Accsum + 19.99
End Sub

Private Sub Units_Click() 'buys accessory and adds cost to total
picResultsaccessories.Print "total Percussion Unit", FormatCurrency(299, 2)
Accsum = Accsum + 299
End Sub

Private Sub Sheild_Click() 'buys accessory and adds cost to total
picResultsaccessories.Print "Remo Drum Sheild", FormatCurrency(250, 2)
Accsum = Accsum + 250
End Sub

Private Sub Heads_Click()   'buys accessory and adds cost to total
picResultsaccessories.Print "Tama Drum Heads", FormatCurrency(45, 2)
Accsum = Accsum + 45
End Sub

Private Sub Sticks_Click()
Number = InputBox("enter the number of sticks you would like to order in multiples of 20", "Place Order")
Select Case Number                'input number of sticks you would like
Case Is = 20
    picResultsaccessories.Print "Firth Drum Sticks", FormatCurrency(30, 2)
    Accsum = Accsum + 30          'buys 20 sticks and adds cost to total
Case 40
    picResultsaccessories.Print "Firth Drum Sticks", FormatCurrency(60, 2)
    Accsum = Accsum + 60          'buys 40 sticks and adds cost to total
Case 60
    picResultsaccessories.Print "Firth Drum Sticks", FormatCurrency(90, 2)
    Accsum = Accsum + 90          'buys 60 sticks and adds cost to total
Case Else
    MsgBox "Please enter a multiple of 20 between 20 and 60", , "Invalid Request"
End Select                        'if a wrong quantity is typed display invalid message
End Sub

Private Sub Pedals_Click()   'buys accessory and adds cost to total
picResultsaccessories.Print "Cobra Double-Bass Pedal", FormatCurrency(129.99, 2)
Accsum = Accsum + 129.99
End Sub

Private Sub Thrones_Click()   'buys accessory and adds cost to total
picResultsaccessories.Print "Drum Thrones", FormatCurrency(69.5, 2)
Accsum = Accsum + 69.5
End Sub
