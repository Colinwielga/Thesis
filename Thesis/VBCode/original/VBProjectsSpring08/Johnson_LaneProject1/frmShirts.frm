VERSION 5.00
Begin VB.Form frmShirts 
   Caption         =   "Shirts"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   Picture         =   "frmShirts.frx":0000
   ScaleHeight     =   8910
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   9000
      TabIndex        =   26
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   495
      Left            =   4200
      TabIndex        =   17
      Top             =   2280
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      Height          =   1815
      Left            =   2400
      ScaleHeight     =   1755
      ScaleWidth      =   4875
      TabIndex        =   16
      Top             =   240
      Width           =   4935
   End
   Begin VB.CommandButton cmdInfo8 
      Caption         =   "Info"
      Height          =   375
      Left            =   7920
      TabIndex        =   15
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo7 
      Caption         =   "Info"
      Height          =   375
      Left            =   7920
      TabIndex        =   14
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo6 
      Caption         =   "Info"
      Height          =   375
      Left            =   7800
      TabIndex        =   13
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo5 
      Caption         =   "Info"
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo4 
      Caption         =   "Info"
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo3 
      Caption         =   "Info"
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo2 
      Caption         =   "Info"
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo1 
      Caption         =   "Info"
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   2280
      Width           =   975
   End
   Begin VB.PictureBox Picture8 
      Height          =   1575
      Left            =   7560
      Picture         =   "frmShirts.frx":10C2A
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   5160
      Width           =   1575
   End
   Begin VB.PictureBox Picture7 
      Height          =   1575
      Left            =   600
      Picture         =   "frmShirts.frx":116BB
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   5280
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   1575
      Left            =   7680
      Picture         =   "frmShirts.frx":12190
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   2760
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   1575
      Left            =   600
      Picture         =   "frmShirts.frx":12B24
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   3960
      Picture         =   "frmShirts.frx":13433
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   5880
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   7560
      Picture         =   "frmShirts.frx":13E6E
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   3960
      Picture         =   "frmShirts.frx":1456D
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   3480
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   600
      Picture         =   "frmShirts.frx":156A2
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "8"
      Height          =   255
      Left            =   9240
      TabIndex        =   25
      Top             =   5160
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "7"
      Height          =   255
      Left            =   9360
      TabIndex        =   24
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "6"
      Height          =   255
      Left            =   9240
      TabIndex        =   23
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "5"
      Height          =   255
      Left            =   3720
      TabIndex        =   22
      Top             =   5880
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "4"
      Height          =   255
      Left            =   3720
      TabIndex        =   21
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "3"
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   5280
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "2"
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   600
      Width           =   135
   End
End
Attribute VB_Name = "frmShirts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmShirts
'Author: Sean Johnson and Nick Lane
'Date Written: Sunday March 16th, 2007
'Objective of form: this is the women shirt form and from this form the
'                   user can find out information about the various items available,
'                   buy an item, and view the price of the item along with a running
'                   total of other items.

Dim Shirt(1 To 20) As String
Dim Prices(1 To 20) As Single
Dim ctr As Integer, j As Integer, c As String
Dim found As Boolean, n As Integer, i As Integer
Option Explicit

Private Sub cmdBack_Click()
'this button will hide this form and show the previous form
frmShirts.Hide
frmWomenApparel.Show
End Sub

'this button will allow the user to purchase an item
Private Sub cmdBuy_Click()
found = False

ctr = 0

'opens the array file to be read into a parallel array
Open App.Path & "\ShirtArray.txt" For Input As #1

'this loop will read the file into a parallel array which will be used when users selects an items
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Shirt(ctr), Prices(ctr)
Loop
Close #1

'allows user to insert the number of the corresponding item he/she wishes to purchase or add to cart.
n = InputBox("Please enter the Number of the corresponding Shoe You wish to purchase")

'this for loop will match the corresponding number above with the corresponding items in the array
For j = 1 To ctr
    If n = j Then
        found = True
        picResults.Print Shirt(j), Tab(25); Prices(j)
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
MsgBox "The Nike Women's Tempo Top is perfect for everyday runs and race day. Mesh insets give a flash of color while providing ventilation in high heat zones. With a V-neck for a youthful edge, this running top is a sure success. Swoosh design trademark embroidered at left chest. Dri-FIT® 100% polyester mesh. Imported", , "Nike Revolution Women's Long Sleeved Tempo Top "
End Sub

Private Sub cmdInfo2_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Women's Tempo Top is perfect for everyday runs and race day. Mesh insets give a flash of color while providing ventilation in high heat zones. With a V-neck for a youthful edge, this running top is a sure success. Swoosh design trademark embroidered at left chest. Dri-FIT® 100% polyester mesh. Imported.", , "Nike Women's L/S Tempo Top"
End Sub

Private Sub cmdInfo3_Click()    'allows the user to view the specific information on the item
MsgBox "Take the trails or race the road in the Nike Women's Adventure Half-Zip Top. Made of 92% Dri-FIT® polyester/8% spandex with mesh insets made of 80% polyester/20% spandex. This versatile top has performance features such as a dropped back hem, a mesh back that wraps to the front, and strategic reflectivity for safety. A Swoosh design trademark is located at the upper left collar. Imported.", , "Nike Women's Short Sleeved Adventure Top"
End Sub

Private Sub cmdInfo4_Click()    'allows the user to view the specific information on the item
MsgBox "Love to work out and still feel pretty? The new Nike Women's Essential Stripe Top offers a distinctly feminine silhouette thanks to cap sleeves, cotton/spandex blend to form-fit and added body length. Stripe detailing adds some spice to this top you'll love to work out in! Nike logo at left chest. 95% cotton/5% spandex. Imported.", , "Nike Women's Essential Stripe Top"
End Sub

Private Sub cmdInfo5_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Iced Over Tee is slightly fitted with distressed graphics on the center front. A versatile addition to any sporty wardrobe. 65% polyester/35% cotton plain jersey. Imported.", , "Nike Women's Iced Over Track Short Sleeve Tee"
End Sub

Private Sub cmdInfo6_Click()    'allows the user to view the specific information on the item
MsgBox "Why live with demanding, tight, or overly baggy tees? The Nike Women's Freeze Frame Relaxed Tee is just the right balance of fit and freedom that will have you moving comfortably. The Freeze Frame tee features a slightly relaxed cut with a center front foil application Swoosh design trademark. 100% cotton plain jersey.", , "Nike Women's Freeze Frame Relaxed Tee"
End Sub

Private Sub cmdInfo7_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Women's Pro Vent Camo running top is a loose-knit top with contrast stitching on raglan seams. Swoosh design trademark heat transfer at left chest. Nike FIT trademark heat transfer at lower right front. Dri-FIT® 69% polyester/17% nylon/14% spandex plain jersey. Imported.", , "Nike Women's Pro Vent ShortSleeve Camo Top"
End Sub

Private Sub cmdInfo8_Click()    'allows the user to view the specific information on the item
MsgBox "Get a jumpstart on a new fitness routine or simply enjoy a casual evening walk in the Nike Spring Sport Tee. This slightly-fitted tee features Dri-FIT® materials for the ultimate in comfort and feminine appeal. 65% cotton (5% organic)/35% polyester plated jersey. Imported", , "Nike Women's Spring Sport Tee"
End Sub
