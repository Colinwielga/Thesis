VERSION 5.00
Begin VB.Form frmWomenAccessories 
   Caption         =   "Women's Accessories"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   Picture         =   "frmWomenAccessories.frx":0000
   ScaleHeight     =   9015
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   6480
      TabIndex        =   38
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   495
      Left            =   3720
      TabIndex        =   25
      Top             =   5040
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      Height          =   2535
      Left            =   3240
      ScaleHeight     =   2475
      ScaleWidth      =   5115
      TabIndex        =   24
      Top             =   2400
      Width           =   5175
   End
   Begin VB.CommandButton cmdInfo12 
      Caption         =   "Info"
      Height          =   255
      Left            =   8400
      TabIndex        =   23
      Top             =   8040
      Width           =   855
   End
   Begin VB.CommandButton cmdInfo11 
      Caption         =   "Info"
      Height          =   255
      Left            =   6000
      TabIndex        =   22
      Top             =   8040
      Width           =   855
   End
   Begin VB.CommandButton cmdInfo10 
      Caption         =   "Info"
      Height          =   255
      Left            =   3480
      TabIndex        =   21
      Top             =   8040
      Width           =   855
   End
   Begin VB.CommandButton cmdInfo9 
      Caption         =   "Info"
      Height          =   255
      Left            =   960
      TabIndex        =   20
      Top             =   8040
      Width           =   855
   End
   Begin VB.CommandButton cmdInfo8 
      Caption         =   "Info"
      Height          =   255
      Left            =   9240
      TabIndex        =   19
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton cmdInfo7 
      Caption         =   "Info"
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton cmdInfo6 
      Caption         =   "Info"
      Height          =   255
      Left            =   9240
      TabIndex        =   17
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdInfo5 
      Caption         =   "Info"
      Height          =   255
      Left            =   1320
      TabIndex        =   16
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdInfo4 
      Caption         =   "Info"
      Height          =   255
      Left            =   8280
      TabIndex        =   15
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdInfo3 
      Caption         =   "Info"
      Height          =   255
      Left            =   6000
      TabIndex        =   14
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdInfo2 
      Caption         =   "Info"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdInfo1 
      Caption         =   "Info"
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   1800
      Width           =   855
   End
   Begin VB.PictureBox Picture12 
      Height          =   1575
      Left            =   8040
      Picture         =   "frmWomenAccessories.frx":240042
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   11
      Top             =   6360
      Width           =   1575
   End
   Begin VB.PictureBox Picture11 
      Height          =   1575
      Left            =   5640
      Picture         =   "frmWomenAccessories.frx":2407F3
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   10
      Top             =   6360
      Width           =   1575
   End
   Begin VB.PictureBox Picture10 
      Height          =   1575
      Left            =   600
      Picture         =   "frmWomenAccessories.frx":240CD7
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   9
      Top             =   6360
      Width           =   1575
   End
   Begin VB.PictureBox Picture9 
      Height          =   1575
      Left            =   960
      Picture         =   "frmWomenAccessories.frx":241538
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   4200
      Width           =   1575
   End
   Begin VB.PictureBox Picture8 
      Height          =   1575
      Left            =   7920
      Picture         =   "frmWomenAccessories.frx":241F64
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture7 
      Height          =   1575
      Left            =   5640
      Picture         =   "frmWomenAccessories.frx":2429ED
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   1575
      Left            =   3120
      Picture         =   "frmWomenAccessories.frx":2436A2
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   6360
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   1575
      Left            =   960
      Picture         =   "frmWomenAccessories.frx":243C08
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   8880
      Picture         =   "frmWomenAccessories.frx":24422C
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   3120
      Picture         =   "frmWomenAccessories.frx":2450E9
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   600
      Picture         =   "frmWomenAccessories.frx":245756
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   8880
      Picture         =   "frmWomenAccessories.frx":24637D
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "12"
      Height          =   255
      Left            =   7680
      TabIndex        =   37
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label Label11 
      Caption         =   "11"
      Height          =   255
      Left            =   5280
      TabIndex        =   36
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "10"
      Height          =   255
      Left            =   2760
      TabIndex        =   35
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "9"
      Height          =   255
      Left            =   360
      TabIndex        =   34
      Top             =   6360
      Width           =   135
   End
   Begin VB.Label Label8 
      Caption         =   "8"
      Height          =   255
      Left            =   8640
      TabIndex        =   33
      Top             =   4200
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "7"
      Height          =   255
      Left            =   720
      TabIndex        =   32
      Top             =   4200
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "6"
      Height          =   255
      Left            =   8640
      TabIndex        =   31
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "5"
      Height          =   255
      Left            =   720
      TabIndex        =   30
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "4"
      Height          =   255
      Left            =   7680
      TabIndex        =   29
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "3"
      Height          =   255
      Left            =   5400
      TabIndex        =   28
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "2"
      Height          =   255
      Left            =   2880
      TabIndex        =   27
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Left            =   360
      TabIndex        =   26
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "frmWomenAccessories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmWomenAccessories
'Author: Sean Johnson and Nick Lane
'Date Written: Sunday March 16th, 2007
'Objective of form: this is the women accessories form and from this form the
'                   user can find out information about the various items available,
'                   buy an item, and view the price of the item along with a running
'                   total of other items.

Dim Accessories(1 To 20) As String
Dim Prices(1 To 20) As Single
Dim ctr As Integer, j As Integer, c As String
Dim found As Boolean, n As Integer, i As Integer
Option Explicit

Private Sub cmdBack_Click()
'this button will hide this form and show the previous form
frmWomenAccessories.Hide
frmWomen.Show
End Sub

'this button will allow the user to purchase an item
Private Sub cmdBuy_Click()
found = False

ctr = 0

'opens the array file to be read into a parallel array
Open App.Path & "\WomAccessories.txt" For Input As #1

'this loop will read the file into a parallel array which will be used when users selects an items
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Accessories(ctr), Prices(ctr)
Loop
Close #1

'allows user to insert the number of the corresponding item he/she wishes to purchase or add to cart.
n = InputBox("Please enter the Number of the corresponding Shoe You wish to purchase")

'this for loop will match the corresponding number above with the corresponding items in the array
For j = 1 To ctr
    If n = j Then
        found = True
        picResults.Print Accessories(j), Tab(25); Prices(j)
    End If
Next j
    If n > j Then   'this loops gives an error message of the user enters a number that doesnt correspond with the labeled items on the form
        MsgBox "Oooops! You have Entered an invalid Number. Please enter a valid number"
    End If
    
'this loop will keep the running total of items and make it viewable to the users
For i = 1 To ctr
    If n = i Then
        found = True
        sum = sum + Prices(i)
        picResults.Print Tab(25); Tab(50); sum  'prints the users running total
    End If
         
Next i
End Sub

Private Sub cmdInfo1_Click()    'allows the user to view the specific information on the item
MsgBox "Load up your equipment and practice essentials with the Nike Women's Court Classics Medium Duffle. Synthetic leather construction features nylon webbing and custom hardware for a distinctly sporty look. Zippered main compartment and various zip-closure pockets inside and outside provide you with all the storage options you need to keep valuables and small items safe.", , "Nike Women's Court Classics Duffle"
End Sub

Private Sub cmdInfo10_Click()   'allows the user to view the specific information on the item
MsgBox "Nike Siege2 Sunglasses feature an interchange lens system that lets you quickly replace lens tint to match any condition. Adjustable, secure wrap temples grip the back-of-head for stability and comfort. Adjustable, ventilated nose bridge adds customizable fit, reduced fogging and better grip. Velocity cut, flying lens offers reduced fogging, consistent vision.", , "Nike Siege2 Sunglasses"
End Sub

Private Sub cmdInfo11_Click()   'allows the user to view the specific information on the item
MsgBox "With moisture-wicking Dri-FIT® fabric, the Nike Shox Lightweight No-Show running sock keeps feet dry and cool. An anatomical left and right foot fit provides superior fit and cushioning, while the arch support gives a snug, secure fit that is enhanced by the Y-heel pocket. The reinforced heel and toe enhances durability in high-wear areas. Dri-FIT® 97% nylon/3% spandex", , "Nike Women's Shox LightWeight No Show Socks"
End Sub

Private Sub cmdInfo12_Click()   'allows the user to view the specific information on the item
MsgBox "Nike Lacrosse Bicep Bands are made of 72% cotton/10% nylon/10% rubber/8% Lycra® spandex terry with an embroidered Lacrosse logo. ", , "Nike Lacrosse Bicep Bands"
End Sub
    
Private Sub cmdInfo2_Click()    'allows the user to view the specific information on the item
MsgBox "Opt for cushiony-soft comfort during workouts and everyday wear. The Nike Women's Low-Cut Sock features a full-cushion terry foot with arch support for comfort, cushioning and shock absorption. Additional details include a hemmed, double-welt top, a knit-in Swoosh design trademark and reinforced heel and toe for enhanced durability in high-wear areas. This Nike sport sock comes in a package of three. 74% cotton/24% polyester/1% spandex/1% other. Imported.", , "Nike Women's 3 Pack Anklet Sock"
End Sub

Private Sub cmdInfo3_Click()    'allows the user to view the specific information on the item
MsgBox "43-lap memory, two-segment interval timer. Time, date, two time zones, alarm. Oversized display, stainless steel bezel, mineral glass crystal, Nike Electrolite, polyurethane strap. Water-resistant to 50 meters.", , " Nike Imara Run "
End Sub

Private Sub cmdInfo4_Click()    'allows the user to view the specific information on the item
MsgBox "Inspired by one of the greatest golfers in the world, Tiger Woods. The Nike Perforated Leather Glove Belt is made of soft, supple 100% glove leather with a simple, elegant buckle and a Tiger Woods signature logo metal inset on belt strap.", , "Nike Perforated Glove Leather Belt"
End Sub

Private Sub cmdInfo5_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Skylon EXP sunglasses feature Max Lens technology for distortion-free vision at all angles of view. Interchange Lens System offers multiple lens options that permit maximum sport performance in all light conditions. Ventilated nose bridge improves airflow for reduced slippage and fogging. Secure wrap temples grip the back of the head for motion stability. Polycarbonate lenses provide scratch- and impact-resistant protection. 100% UVA and UVB protection.", , "Nike Skylon EXP Sunglasses"
End Sub

Private Sub cmdInfo6_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike B2.9 Gymsack has a drawstring bag closure that provides easy access to contents. Screenprinted Swoosh. 210 denier nylon/75 denier polyester dobby weave", , "Nike B2.9 Gymsack"
End Sub

Private Sub cmdInfo7_Click()    'allows the user to view the specific information on the item
MsgBox "Inspired by one of the greatest golfers in the world, Tiger Woods. The Nike Bridle Leather Belt features a 100% leather strap with a fashion-forward, epoxy-filled, signature buckle featuring Tiger Woods' initials.", , "Nike Bridle Leather Belt"
End Sub

Private Sub cmdInfo8_Click()    'allows the user to view the specific information on the item
MsgBox "Time your workouts with the Nike Triax Mia, the watch designed for women. Features include a small S-shaped design to curve around your wrist and keep the watch in place. Functions include 50-lap chronograph, two-segment interval timer, two alarms, two time zones, date, and one-touch backlighting. A solid, hardened aluminum case, polyurethane strap, and mineral glass crystal ensure durability for the long run. Water resistant to 50 meters.", , "Nike Women's Triax Mia"
End Sub

Private Sub cmdInfo9_Click()    'allows the user to view the specific information on the item
MsgBox "Nike Thermal Running Gloves are made with an insulted fleece fabric that holds in body heat to help keep you warm on cold days. Key pocket provides easy access and secure storage. Reflective pattern adds enhanced visibility. 98% polyester/2% spandex. Imported.", , "Nike Thermal Run Gloves"
End Sub
