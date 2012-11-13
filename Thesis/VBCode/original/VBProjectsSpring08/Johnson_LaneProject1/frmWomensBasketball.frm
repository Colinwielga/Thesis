VERSION 5.00
Begin VB.Form frmWomensBasketball 
   Caption         =   "Women's Basketball Shoes"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   Picture         =   "frmWomensBasketball.frx":0000
   ScaleHeight     =   8655
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   735
      Left            =   9720
      TabIndex        =   29
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdInfo9 
      Caption         =   "Info"
      Height          =   375
      Left            =   7440
      TabIndex        =   19
      Top             =   7200
      Width           =   735
   End
   Begin VB.CommandButton cmdInfo8 
      Caption         =   "Info"
      Height          =   375
      Left            =   5160
      TabIndex        =   18
      Top             =   6840
      Width           =   735
   End
   Begin VB.CommandButton cmdInfo7 
      Caption         =   "Info"
      Height          =   375
      Left            =   3120
      TabIndex        =   17
      Top             =   7200
      Width           =   735
   End
   Begin VB.CommandButton cmdInfo6 
      Caption         =   "Info"
      Height          =   375
      Left            =   960
      TabIndex        =   16
      Top             =   6840
      Width           =   735
   End
   Begin VB.CommandButton cmdInfo5 
      Caption         =   "Info"
      Height          =   375
      Left            =   840
      TabIndex        =   15
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdInfo4 
      Caption         =   "Info"
      Height          =   375
      Left            =   7440
      TabIndex        =   14
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdInfo3 
      Caption         =   "Info"
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdInfo2 
      Caption         =   "Info"
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdInfo1 
      Caption         =   "Info"
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   495
      Left            =   8880
      TabIndex        =   10
      Top             =   4920
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      Height          =   2055
      Left            =   3960
      ScaleHeight     =   1995
      ScaleWidth      =   6075
      TabIndex        =   9
      Top             =   2760
      Width           =   6135
   End
   Begin VB.PictureBox Picture9 
      Height          =   1575
      Left            =   2640
      Picture         =   "frmWomensBasketball.frx":240042
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox Picture8 
      Height          =   1575
      Left            =   6960
      Picture         =   "frmWomensBasketball.frx":241881
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   5520
      Width           =   1575
   End
   Begin VB.PictureBox Picture7 
      Height          =   1575
      Left            =   4800
      Picture         =   "frmWomensBasketball.frx":2421AB
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   5160
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   1575
      Left            =   480
      Picture         =   "frmWomensBasketball.frx":242D8A
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   1575
      Left            =   6960
      Picture         =   "frmWomensBasketball.frx":243750
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   4800
      Picture         =   "frmWomensBasketball.frx":2442FA
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   480
      Picture         =   "frmWomensBasketball.frx":244F45
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   5160
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   2640
      Picture         =   "frmWomensBasketball.frx":24563B
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   5520
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   480
      Picture         =   "frmWomensBasketball.frx":246FDE
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "9"
      Height          =   255
      Left            =   6840
      TabIndex        =   28
      Top             =   5520
      Width           =   135
   End
   Begin VB.Label Label8 
      Caption         =   "8"
      Height          =   255
      Left            =   4680
      TabIndex        =   27
      Top             =   5160
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "7"
      Height          =   255
      Left            =   2520
      TabIndex        =   26
      Top             =   5520
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "6"
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   5160
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "5"
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "4"
      Height          =   255
      Left            =   6840
      TabIndex        =   23
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "3"
      Height          =   255
      Left            =   4680
      TabIndex        =   22
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "2"
      Height          =   255
      Left            =   2520
      TabIndex        =   21
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   600
      Width           =   135
   End
End
Attribute VB_Name = "frmWomensBasketball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmWomenBasketball
'Author: Sean Johnson and Nick Lane
'Date Written: Sunday March 16th, 2007
'Objective of form: this is the women basketball shoes form and from this form the
'                   user can find out information about the various items available,
'                   buy an item, and view the price of the item along with a running
'                   total of other items.

Dim Basketball(1 To 20) As String
Dim Prices(1 To 20) As Single
Dim ctr As Integer, j As Integer, c As String
Dim found As Boolean, n As Integer, i As Integer

Option Explicit

Private Sub cmdBack_Click()
'this button will hide this form and show the previous form
frmWomensBasketball.Hide
frmWomensShoes.Show
End Sub

'this button will allow the user to purchase an item
Private Sub cmdBuy_Click()
found = False

ctr = 0

'opens the array file to be read into a parallel array
Open App.Path & "\BasketballArray.txt" For Input As #1

'this loop will read the file into a parallel array which will be used when users selects an items
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Basketball(ctr), Prices(ctr)
Loop
Close #1

'allows user to insert the number of the corresponding item he/she wishes to purchase or add to cart.
n = InputBox("Please enter the Number of the corresponding Shoe You wish to purchase")

'this for loop will match the corresponding number above with the corresponding items in the array
For j = 1 To ctr
    If n = j Then
        found = True
        picResults.Print Basketball(j), Tab(25); Prices(j)
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
MsgBox "The Nike Women's Zoom Huarache Elite basketball shoe is inspired by WNBA legends Sheryl Swoopes, Swin Cash and Sue Bird. Lightweight full-grain leather upper provides anatomical fit with an ankle strap for superior lockdown. Phylon™ midsole with natural motion principles applied to the forefoot to increase flexibility. Zoom Air™ unit in the forefoot and heel. Modified outsole pattern with herringbone in the forefoot and heel adds optimal traction. Wt. 13.2 oz.", , "Nike Women's Zoom Huarache Elite"
End Sub

Private Sub cmdInfo2_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Air Finesse Uptempo is a lightweight, supportive basketball shoe for the young player who is looking for the best in signature level value, attitude and style. Synthetic leather upper. Support strap adds lockdown fit. Integrated midfoot TPU Nike Chassis system provides lateral support. Internal Phylon™ midsole with visible heel Air-Sole® unit surrounded by transparent TPU clip. Solid rubber outsole with herringbone pattern offers maximum on court traction.", , "Nike Women's Air Finesse Uptempo"
End Sub

Private Sub cmdInfo3_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Vis-Air Elite basketball shoe features a full-grain leather upper with molded overlays for maximum support. Lightweight and durable PU midsole. Fully exposed large-volume Air-Sole® unit in the heel provides cushioning. Solid rubber outsole with modified herringbone pattern offers enhanced traction. TPU shank adds lateral stability. Wt. 15.2 oz.", , "Nike Women's VIS Air Elite"
End Sub

Private Sub cmdInfo4_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Air Huarache Excel is a lightweight, low-to-the-ground basketball shoe with increased forefoot flexibility and an extremely responsive ride to meet the demands of the ultimate female speed player. Inspired by WNBA legends — Sheryl Swoopes, Swin Cash and Sue Bird. Lightweight synthetic leather upper provides anatomical fit. Proprioceptive ankle strap with integrated medial chassis adds superior lockdown and stability. Lateral TPU heel plate gives stabilized heel movement. Perforated lateral window offers enhanced breathability. Lightweight Phylon™ midsole with natural motion principles applied to forefoot region for increased flexibility. Regional responsive Zoom Air™ units in the forefoot and heel. Modified outsole pattern with herringbone in the forefoot and heel for optimal traction. Wt. 13.4 oz.", , "Nike Women's Air Huarache Excel"
End Sub

Private Sub cmdInfo5_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Air Flight Dime Dropper is a fast basketball shoe for players looking to take their game to the next level. Double-lasted to get the player lower to the court for a great feel. Phylon™ midsole with a heel Air-Sole® unit provides great cushioning. Traditional herringbone outsole offers optimal traction on the court. Outrigger adds lateral stability. Wt. 12.2 oz.", , "Nike Women's Air Flight Dime Dropper"
End Sub

Private Sub cmdInfo6_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Air Total Package is a lightweight, comfortable, low-cut basketball shoe for the player seeking cutting-edge performance and style. Leather and synthetic leather upper with speed lacing for convenience. Full-length Phylon™ midsole with encapsulated Air-Sole™ unit provides comfort and cushioning. Solid rubber outsole with herringbone tread delivers excellent traction. Wt. 12.6 oz.", , "Nike Women's Air Total Package Low"
End Sub
    
Private Sub cmdInfo7_Click()    'allows the user to view the specific information on the item
MsgBox "The arrogance of Flight meets the attitude of Diana Taurasi. The Nike Shox Elite basketball shoe is protective, durable and lightweight for today's baller. Full-grain leather upper provides anatomical fit with an ankle strap for suprior lockdown. Full-length Phylon™ midsole with Nike Shox™ technology in the heel. Forefoot Zoom Air™ unit. Podular herringbone outsole delivers maximum indoor/outdoor traction. Wt. 15.0 oz.", , "Nike Women's Shox Elite"
End Sub

Private Sub cmdInfo8_Click()    'allows the user to view the specific information on the item
MsgBox "Superior craftsmanship meets cutting edge performance in a modern, innovative flight shoe tuned to the specific needs of the elite female basketball player. The Nike Shox DT is inspired by the magnetic personality and play of the future of women's game — Diana Taurasi. Rich synthetic leather upper with technical high performance materials. Molded anatomical heel bucket cradles the foot for enhanced ankle and heel stability. Integrated synthetic harness delivers midfoot lockdown. Full internal innersleeve provides seamless, comfortable fit. Double-lasted Phylon™ midsole. Nike Shox™ technology in the heel combined with Zoom Air™ unit in the forefoot. Solid rubber outsole with herringbone pattern adds maximum traction. Wt. 15.8 oz.", , "Nike Women's Shox DT"
End Sub
    
Private Sub cmdInfo9_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike LeBron III is a lightweight, comfortable and stable basketball shoe designed for the player who can do it all on the court. Inspired by the player who is the future of the game - LeBron James. Lightweight and breathable synthetic upper. Leather support straps adds lateral stability. Lightweight Phylon™ midsole. Durable, non-marking solid rubber outsole with maximum traction modified herringbone pattern. Wt. 15.2 oz.", , "Nike Women's Zoom LeBron III Low "
End Sub
