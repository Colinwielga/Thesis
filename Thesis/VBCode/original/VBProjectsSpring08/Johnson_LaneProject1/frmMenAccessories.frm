VERSION 5.00
Begin VB.Form frmMenAccessories 
   Caption         =   "Men Accessories"
   ClientHeight    =   10890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   Picture         =   "frmMenAccessories.frx":0000
   ScaleHeight     =   10890
   ScaleWidth      =   13125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   495
      Left            =   5160
      TabIndex        =   47
      Top             =   9960
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      Height          =   1695
      Left            =   480
      ScaleHeight     =   1635
      ScaleWidth      =   6555
      TabIndex        =   46
      Top             =   8040
      Width           =   6615
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   375
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   1695
   End
   Begin VB.PictureBox Picture15 
      Height          =   1575
      Left            =   10440
      Picture         =   "frmMenAccessories.frx":18004E
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   29
      Top             =   7920
      Width           =   1575
   End
   Begin VB.PictureBox Picture14 
      Height          =   1575
      Left            =   8160
      Picture         =   "frmMenAccessories.frx":180CD1
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   28
      Top             =   8040
      Width           =   1575
   End
   Begin VB.PictureBox Picture13 
      Height          =   1575
      Left            =   10440
      Picture         =   "frmMenAccessories.frx":181639
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   27
      Top             =   5760
      Width           =   1575
   End
   Begin VB.PictureBox Picture12 
      Height          =   1575
      Left            =   8040
      Picture         =   "frmMenAccessories.frx":182B0B
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   26
      Top             =   5760
      Width           =   1575
   End
   Begin VB.PictureBox Picture11 
      Height          =   1575
      Left            =   5640
      Picture         =   "frmMenAccessories.frx":18360E
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   25
      Top             =   5760
      Width           =   1575
   End
   Begin VB.PictureBox Picture10 
      Height          =   1575
      Left            =   3000
      Picture         =   "frmMenAccessories.frx":183DB7
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   24
      Top             =   5640
      Width           =   1575
   End
   Begin VB.PictureBox Picture9 
      Height          =   1575
      Left            =   600
      Picture         =   "frmMenAccessories.frx":184A59
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   23
      Top             =   5640
      Width           =   1575
   End
   Begin VB.PictureBox Picture8 
      Height          =   1575
      Left            =   8040
      Picture         =   "frmMenAccessories.frx":18568C
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   22
      Top             =   3120
      Width           =   1575
   End
   Begin VB.PictureBox Picture7 
      Height          =   1575
      Left            =   5520
      Picture         =   "frmMenAccessories.frx":186200
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   21
      Top             =   3120
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   1575
      Left            =   5520
      Picture         =   "frmMenAccessories.frx":186DD9
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   20
      Top             =   720
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   1575
      Left            =   2760
      Picture         =   "frmMenAccessories.frx":1876A5
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   19
      Top             =   3000
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   480
      Picture         =   "frmMenAccessories.frx":188FEC
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   18
      Top             =   3000
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   7920
      Picture         =   "frmMenAccessories.frx":189DF3
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   17
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdInfo3 
      Caption         =   "Info"
      Height          =   255
      Left            =   7920
      TabIndex        =   16
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdInfo4 
      Caption         =   "Info"
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdInfo5 
      Caption         =   "Info"
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdInfo6 
      Caption         =   "Info"
      Height          =   255
      Left            =   5640
      TabIndex        =   13
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdinfo7 
      Caption         =   "Info"
      Height          =   255
      Left            =   5640
      TabIndex        =   12
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdInfo8 
      Caption         =   "Info"
      Height          =   255
      Left            =   8160
      TabIndex        =   11
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdInfo9 
      Caption         =   "Info"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton cmdInfo10 
      Caption         =   "Info"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton cmdInfo11 
      Caption         =   "Info"
      Height          =   255
      Left            =   5760
      TabIndex        =   8
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdInfo12 
      Caption         =   "Info"
      Height          =   255
      Left            =   8160
      TabIndex        =   7
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdinfo13 
      Caption         =   "Info"
      Height          =   255
      Left            =   10560
      TabIndex        =   6
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdInfo14 
      Caption         =   "Info"
      Height          =   255
      Left            =   8280
      TabIndex        =   5
      Top             =   9720
      Width           =   1335
   End
   Begin VB.CommandButton cmdInfo15 
      Caption         =   "Info"
      Height          =   255
      Left            =   10560
      TabIndex        =   4
      Top             =   9600
      Width           =   1335
   End
   Begin VB.CommandButton cmdInfo2 
      Caption         =   "Info"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   2880
      Picture         =   "frmMenAccessories.frx":18AD5E
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Info"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   480
      Picture         =   "frmMenAccessories.frx":18B2C4
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "15"
      Height          =   255
      Index           =   14
      Left            =   10080
      TabIndex        =   45
      Top             =   8040
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "14"
      Height          =   255
      Index           =   13
      Left            =   7800
      TabIndex        =   44
      Top             =   8160
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "13"
      Height          =   255
      Index           =   12
      Left            =   10080
      TabIndex        =   43
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "12"
      Height          =   255
      Index           =   11
      Left            =   7680
      TabIndex        =   42
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "11"
      Height          =   255
      Index           =   10
      Left            =   5280
      TabIndex        =   41
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "10"
      Height          =   255
      Index           =   9
      Left            =   2640
      TabIndex        =   40
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "9"
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   39
      Top             =   5640
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "8"
      Height          =   255
      Index           =   7
      Left            =   7800
      TabIndex        =   38
      Top             =   3240
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "7"
      Height          =   255
      Index           =   6
      Left            =   5280
      TabIndex        =   37
      Top             =   3240
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "6"
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   36
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "5"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   35
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   7680
      TabIndex        =   34
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   33
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   32
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   31
      Top             =   840
      Width           =   135
   End
End
Attribute VB_Name = "frmMenAccessories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Project name: Nike Town
'Form name: frmMenAccessories
'Author: Sean Johnson and Nick Lane
'Date Written: Saturday March 15th, 2007
'Objective of form: this forms is the men's accessories form and from this form the
'                   user can find out information about the various items available,
'                   buy an item, and view the price of the item along with a running
'                   total of other items.

Option Explicit
Dim Accessory(1 To 20) As String
Dim Prices(1 To 20) As Single
Dim ctr As Integer, j As Integer, c As String
Dim found As Boolean, n As Integer, i As Integer
Private Sub cmdBack_Click()
'this button will hide this form and show the previous form
frmMenAccessories.Hide
frmMen.Show
End Sub

'this button will allow the user to purchase an item
Private Sub cmdBuy_Click()

found = False

ctr = 0

'opens the array file to be read into a parallel array
Open App.Path & "\AccessoriesArray.txt" For Input As #1

'this loop will read the file into a parallel array which will be used when users selects an items
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Accessory(ctr), Prices(ctr)
Loop
Close #1

'allows user to insert the number of the corresponding item he/she wishes to purchase or add to cart.
n = InputBox("Please enter the Number of the corresponding Shoe You wish to purchase")

'this for loop will match the corresponding number above with the corresponding items in the array
For j = 1 To ctr
    If n = j Then
        found = True
        picResults.Print Accessory(j), Tab(25); Prices(j)
    End If
Next j
    If n > j Then 'this loops gives an error message of the user enters a number that doesnt correspond with the labeled items on the form
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

Private Sub cmdInfo_Click() 'allows the user to view the specific information on the item
MsgBox "The Nike Vision Pursue Sunglasses feature Max lens technology that gives you precise visual information at all angles of view. Fying lens has reduced fogging and consistent visual performance. Interchange lens system has multiple lens options which permit maximum sport performance in all light conditions. Ventilated nose bridge improves airflow for reduced slippage and fogging. Adjustable secure wrap temples that grip the back of the head for motion stability. Polycarbonate lenses provide scratch and impact resistant protection. State-of-the-art UV absorbency. 100% UVA and UVB protection.", , "Nike Vision Pursue Sunglasses"
End Sub

Private Sub cmdInfo10_Click()   'allows the user to view the specific information on the item
MsgBox "The Nike Brasilia 4 Duffle is a versatile duffle bag for those on-the-go. Featuring a zippered main compartment and an interior hanging pocket, the Brasilia 4 holds your essential game gear or traveling supplies. Padded, adjustable shoulder straps and convenient carry handles keep the load manageable. Includes an ID label and blank side panel for customization. Made of 210 denier nylon with PVC laminate backing.", , "Nike Brasilia 4 Duffel-Large"
End Sub

Private Sub cmdInfo11_Click()   'allows the user to view the specific information on the item
MsgBox "An athletically inspired belt that's perfect for casual wear. The Nike Web Belt features a synthetic strap with stitch edges to prevent fray and a sleek buckle with embossed signature Swoosh design trademark.", , "Nike Web Belt"
End Sub

Private Sub cmdInfo12_Click()   'allows the user to view the specific information on the item
MsgBox "The Nike Gymsack is a supreme style for carrying your accessories. This 600D polyester bag has a drawstring closure, reflectivity for safety and a screen Nike and Swoosh design trademark. There is an exterior side water bottle pocket and an audio pocket with cord portal so you can keep your tunes with you wherever you go. Imported.", , "Nike Gymsack"
End Sub

Private Sub cmdinfo13_Click()   'allows the user to view the specific information on the item
MsgBox "The Nike Shin Sock III is ideal for young soccer players. All-in-one full sock with shinguard. Ventilated hard-plastic shell and EVA foam padding. Full sock with fitted heel. Built-in foam disks on both sides of the ankle.", , "Nike Shin Sock III"
End Sub

Private Sub cmdInfo14_Click()   'allows the user to view the specific information on the item
MsgBox "Your complete monitoring solution. Download workout data to your computor, upload workout plans and watch settings to your watch. Full featured training software helps you set goals, build training plans, log workouts and analyze the results. You can even customize the display to highlight the information that is important to you. Digital transmission eliminates cross-talk from other wireless devices, target training zones offer programmable pace and HR training zones with audible out-of-zone alert. Accurate heart rate and pace information, 100 lap memory.", , "Nike Triax Elite HRM/SDM"
End Sub

Private Sub cmdInfo15_Click()   'allows the user to view the specific information on the item
MsgBox "Perfect for travel between games, the Nike Team Ball Bag is made of 100% nylon with full-length zipper closure, padded shoulder strap for comfortable carrying, a large side pocket and one small pocket to store smaller items. Holds six official-size basketballs.", , "Nike Team Ball Bag"
End Sub

Private Sub cmdInfo2_Click()    'allows the user to view the specific information on the item
MsgBox "Nike Siege2 Sunglasses feature an interchange lens system that lets you quickly replace lens tint to match any condition. Adjustable, secure wrap temples grip the back-of-head for stability and comfort. Adjustable, ventilated nose bridge adds customizable fit, reduced fogging and better grip. Velocity cut, flying lens offers reduced fogging, consistent vision.", , "Nike Siege2 Sunglasses"
End Sub

Private Sub cmdInfo3_Click()    'allows the user to view the specific information on the item
MsgBox "The 300-lap chronograph on the Nike Triax Vapor 300 is only one of the many features of this technical watch that will keep you training at the top of your game. Additional features include a convertible display and easy-to-read data mode, a five-segment interval timer, two alarms, two time zones, and the date. The pre-curved ergonomic display and band along with a one-touch back light that lasts six seconds when chronograph is activated make for easy reading on the run, while the 100m water resistance and solid aluminum case enhance durability. Nike Swoosh design trademark sits above the face.", , "Nike Triax Vapor 300 Super"
End Sub

Private Sub cmdInfo4_Click()    'allows the user to view the specific information on the item
MsgBox "You'll be watching the time more than usual with the sleek lines and sharp functions of the Nike Sledge Analog watch. A stainless steel case and pre-curved polyurethane strap make the Sledge perfect for everyday wear while its functions are fit to accompany you on your workouts. Features include date and day subdials, luminescent hands and cross-cut metal buttons. Mineral glass crystal. Water resistant to 50 meters.", , "Nike Men's Sledge Analog"
End Sub

Private Sub cmdInfo5_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Triax Speed 300 watch features a 300 lap chronograph and a co-molded polyurethane strap. Convertible display, data mode. Five segment interval timer. Mineral glass crystal. One-touch backlighting. S-shape design curves around wrist. Solid, hardened aluminum case. Battery hatch on back plate. Stainless steel buckle and back plate. Time, date, two alarms, two time zones, target time. 100m water resistance.", , "Nike Triax Speed 300"
End Sub

Private Sub cmdInfo6_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Skylon EXP sunglasses feature Max Lens technology for distortion-free vision at all angles of view. Interchange Lens System offers multiple lens options that permit maximum sport performance in all light conditions. Ventilated nose bridge improves airflow for reduced slippage and fogging. Secure wrap temples grip the back of the head for motion stability. Polycarbonate lenses provide scratch- and impact-resistant protection. 100% UVA and UVB protection.", , "Nike Skylon EXP Sunglasses"
End Sub

Private Sub cmdInfo7_Click()    'allows the user to view the specific information on the item
MsgBox "Keeping time is only the beginning for the Nike Digital Super Chronograph. With 100-hour chronograph and data recall functions, three time alarms and hydration alarm, the Digital Super Chronograph is ready for you and your adventures. The lightweight polymer case and high-contrast LCD display make it ideal for a variety of activities and conditions. Mineral glass crystal and back battery hatch. Water resistant to 100 meters.", , "Nike Men's Digital Super Chronograph"
End Sub

Private Sub cmdInfo8_Click()    'allows the user to view the specific information on the item
MsgBox "An extra large bag with extra large style, the Nike Team Training XL Backpack is made of 600 denier polyester and features a screened Nike design trademark and a screen Swoosh design trademark. A dual zippered main compartment will fit your large items, while the exterior side pocket fits a water bottle and an audio pocket will store your music device with a cord portal for easy access. The shoulder strap is padded and the bag comes with a limited lifetime guarantee.", , "Nike Team Training XL Backpack"
End Sub

Private Sub cmdInfo9_Click()    'allows the user to view the specific information on the item
MsgBox "Inspired by one of the greatest golfers in the world, Tiger Woods. The Nike Italian Luxury Signature Buckle Belt features a 100% leather strap with embossed lines for a classic, elegant look. A large Tiger Woods logo is etched on the buckle.", , "Nike Italian Lux Belt w/ Signature"
End Sub



