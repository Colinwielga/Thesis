VERSION 5.00
Begin VB.Form frmSweats 
   Caption         =   "Sweats"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   Picture         =   "frmSweats.frx":0000
   ScaleHeight     =   9660
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   495
      Left            =   4680
      TabIndex        =   23
      Top             =   6720
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1755
      ScaleWidth      =   7155
      TabIndex        =   22
      Top             =   4680
      Width           =   7215
   End
   Begin VB.CommandButton cmdBmen 
      Caption         =   "Back"
      Height          =   375
      Left            =   9120
      TabIndex        =   14
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Info"
      Height          =   255
      Left            =   7800
      TabIndex        =   13
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   10560
      TabIndex        =   12
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Info"
      Height          =   255
      Left            =   8040
      TabIndex        =   11
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   10560
      TabIndex        =   10
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Info"
      Height          =   255
      Left            =   10440
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   8160
      TabIndex        =   8
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Info"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.PictureBox Picture7 
      Height          =   1575
      Left            =   7680
      Picture         =   "frmSweats.frx":18004E
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   5280
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   1575
      Left            =   10320
      Picture         =   "frmSweats.frx":180A45
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   2760
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   1575
      Left            =   7800
      Picture         =   "frmSweats.frx":181193
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   10320
      Picture         =   "frmSweats.frx":181AA3
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   5280
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   10320
      Picture         =   "frmSweats.frx":18230D
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   7800
      Picture         =   "frmSweats.frx":182995
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   480
      Picture         =   "frmSweats.frx":182FE0
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "2"
      Height          =   255
      Index           =   6
      Left            =   7560
      TabIndex        =   21
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   20
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "7"
      Height          =   255
      Index           =   4
      Left            =   10080
      TabIndex        =   19
      Top             =   5400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "6"
      Height          =   255
      Index           =   3
      Left            =   7440
      TabIndex        =   18
      Top             =   5400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "5"
      Height          =   255
      Index           =   2
      Left            =   9960
      TabIndex        =   17
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "4"
      Height          =   255
      Index           =   1
      Left            =   7560
      TabIndex        =   16
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "3"
      Height          =   255
      Index           =   0
      Left            =   9960
      TabIndex        =   15
      Top             =   480
      Width           =   135
   End
End
Attribute VB_Name = "frmSweats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmSweats
'Author: Sean Johnson and Nick Lane
'Date Written: Saturday March 15th, 2007
'Objective of form: this is the men's sweats form and from this form the
'                   user can find out information about the various items available,
'                   buy an item, and view the price of the item along with a running
'                   total of other items.

Option Explicit
Dim sweats(1 To 20) As String
Dim Prices(1 To 20) As Single
Dim ctr As Integer, j As Integer, c As String
Dim found As Boolean, n As Integer, i As Integer
Private Sub cmdBMen_Click()
'this button will hide this form and show the previous form
frmMenApparel.Show
frmSweats.Hide
End Sub

'this button will allow the user to purchase an item
Private Sub cmdBuy_Click()
found = False

ctr = 0

'opens the array file to be read into a parallel array
Open App.Path & "\sweatsArray.txt" For Input As #1

'this loop will read the file into a parallel array which will be used when users selects an items
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, sweats(ctr), Prices(ctr)
Loop
Close #1

'allows user to insert the number of the corresponding item he/she wishes to purchase or add to cart.
n = InputBox("Please enter the Number of the corresponding Shoe You wish to purchase")

'this for loop will match the corresponding number above with the corresponding items in the array
For j = 1 To ctr
    If n = j Then
        found = True
        picResults.Print sweats(j), Tab(25); Prices(j)
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

Private Sub Command1_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Fleece Pant is a fleece pant with side welt pockets and embroidered Swoosh design trademark at left hip. 80% cotton/20% polyester. Imported.", , "Nike Men's Fleece Pant"
End Sub

Private Sub Command2_Click()    'allows the user to view the specific information on the item
MsgBox "100% woven polyester with contrast piping and embroidered Swoosh design trademark. Mesh lined with zippered pockets. Imported.", , "Nike Men's Solid Classic Woven Pant"
End Sub

Private Sub Command3_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Game Warmup Pant is the official game day tear away pant of the NCAA. Made of 100% polyester with contrast colors on the side panels. Featuring an embriodered team logo on the right leg and an embroidered Swoosh design trademark on the left leg. Imported.", , "Nike Men's Game Warmup Pants-NorthCarolina"
End Sub

Private Sub Command4_Click()    'allows the user to view the specific information on the item
MsgBox "Look the part of a hoops star with the dazzle of the Nike Kobe Diamond Pant! Embossed, shimmery fabric insets shine under the city lights. The Kobe Diamond Pant is a pull-on pant featuring an elastic, drawcord waist and side pockets dressed up with Swoosh and Kobe logos. 100% polyester knit. Imported.", , "Nike Men's Kobe Diamond Pant"
End Sub

Private Sub Command5_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Pro is a relaxed fit training pant with a 12 inch zipper at the bottom leg. Side seam zippered pockets. Swoosh design trademark at the lower back left leg. Dri-FIT® trademark at the lower front left leg. Heat transfer logos. Dri-FIT® 91% polyester/9% spandex. Imported.", , "Nike Men's Pro Touring Pants"
End Sub

Private Sub Command6_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Conquer Game Pant has no competition! This warm-up pant is available in team color with white side insets and drawcord waist. Pant also has side pockets and three ankle snaps to allow versatile hem width. Designed to match Nike Conquer Game Jacket. Embroidered Swoosh design trademark at front left hip. 100% polyester Dri-FIT double knit jacquard. Imported.", , "Nike Men's Conquer Game Pant"
End Sub

Private Sub Command7_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike LeBron Fleece Pant is made of 80% cotton/20% polyester heavyweight fleece with a L23 logo embroidered at the left leg. Imported.", , "Nike Men's Lebron Fleece Pant"
End Sub



