VERSION 5.00
Begin VB.Form frmHeadwear 
   Caption         =   "Headwear"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   Picture         =   "frmHeadwear.frx":0000
   ScaleHeight     =   8790
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   1095
      Left            =   8160
      TabIndex        =   26
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   1095
      Left            =   1200
      TabIndex        =   17
      Top             =   4080
      Width           =   975
   End
   Begin VB.PictureBox picResults 
      Height          =   2055
      Left            =   2280
      ScaleHeight     =   1995
      ScaleWidth      =   5715
      TabIndex        =   16
      Top             =   3600
      Width           =   5775
   End
   Begin VB.CommandButton cmdInfo8 
      Caption         =   "Info"
      Height          =   495
      Left            =   8040
      TabIndex        =   15
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo7 
      Caption         =   "Info"
      Height          =   495
      Left            =   5520
      TabIndex        =   14
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo6 
      Caption         =   "Info"
      Height          =   495
      Left            =   3240
      TabIndex        =   13
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo5 
      Caption         =   "Info"
      Height          =   495
      Left            =   720
      TabIndex        =   12
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo4 
      Caption         =   "Info"
      Height          =   495
      Left            =   8280
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo3 
      Caption         =   "Info"
      Height          =   495
      Left            =   5760
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo2 
      Caption         =   "Info"
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo1 
      Caption         =   "Info"
      Height          =   495
      Left            =   840
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.PictureBox Picture8 
      Height          =   1575
      Left            =   7680
      Picture         =   "frmHeadwear.frx":18004E
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   5880
      Width           =   1575
   End
   Begin VB.PictureBox Picture7 
      Height          =   1575
      Left            =   5280
      Picture         =   "frmHeadwear.frx":180CCC
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   5880
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   1575
      Left            =   2880
      Picture         =   "frmHeadwear.frx":18160A
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   5880
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   1575
      Left            =   480
      Picture         =   "frmHeadwear.frx":182097
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   5880
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   7920
      Picture         =   "frmHeadwear.frx":1826FF
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   5400
      Picture         =   "frmHeadwear.frx":182D3C
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   2880
      Picture         =   "frmHeadwear.frx":183609
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   480
      Picture         =   "frmHeadwear.frx":18411B
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "8"
      Height          =   255
      Left            =   7440
      TabIndex        =   25
      Top             =   5880
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "7"
      Height          =   255
      Left            =   5040
      TabIndex        =   24
      Top             =   5880
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "6"
      Height          =   255
      Left            =   2640
      TabIndex        =   23
      Top             =   5880
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "5"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   5880
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "4"
      Height          =   255
      Left            =   7680
      TabIndex        =   21
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "3"
      Height          =   255
      Left            =   5160
      TabIndex        =   20
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "2"
      Height          =   255
      Left            =   2640
      TabIndex        =   19
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   360
      Width           =   135
   End
End
Attribute VB_Name = "frmHeadwear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmHeadwear
'Author: Sean Johnson and Nick Lane
'Date Written: Saturday March 15th, 2007
'Objective of form: this is the women headwear form and from this form the
'                   user can find out information about the various items available,
'                   buy an item, and view the price of the item along with a running
'                   total of other items.

Dim Headwear(1 To 20) As String
Dim Prices(1 To 20) As Single
Dim ctr As Integer, j As Integer, c As String
Dim found As Boolean, n As Integer, i As Integer
Option Explicit

Private Sub cmdBack_Click()
'this button will hide this form and show the previous form
frmHeadwear.Hide
frmWomenApparel.Show
End Sub

'this button will allow the user to purchase an item
Private Sub cmdBuy_Click()
found = False

ctr = 0

'opens the array file to be read into a parallel array
Open App.Path & "\HeadwearArray.txt" For Input As #1

'this loop will read the file into a parallel array which will be used when users selects an items
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Headwear(ctr), Prices(ctr)
Loop
Close #1

'allows user to insert the number of the corresponding item he/she wishes to purchase or add to cart.
n = InputBox("Please enter the Number of the corresponding Shoe You wish to purchase")

'this for loop will match the corresponding number above with the corresponding items in the array
For j = 1 To ctr
    If n = j Then
        found = True
        picResults.Print Headwear(j), Tab(25); Prices(j)
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

Private Sub cmdInfo1_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Breast Cancer Wool Classic Soccer Cap is made of 100% wool with embroidery. Imported.", , "Nike Bc Wool Classic Soccer Cap "
End Sub

Private Sub cmdInfo2_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Coaches Cap is a six-panel brushed twill cap with heavy buckram structure. Swoosh design trademark embroidery on the left side. Dri-FIT® headband. 100% cotton. One size fits most. Imported.", , "Nike Coaches Cap"
End Sub

Private Sub cmdInfo3_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Women's FIT Reflective Running Hat is a choice cap for your run or jog because of all its great features. Built-in terry sweatband controls moisture along with Dri-FIT® materials and mesh panels for ventilation. All-over reflective print provides 360-degree reflectivity, a must-have for nighttime or pre-dawn runs. Quick adjusting closure can be operated with one hand. Body/bill: Dri-FIT® 100% recycled polyester taffeta. Mesh: Dri-FIT® 100% polyester circular knit mesh. Sweatband: Dri-FIT® 90% polyester/10% spandex. Imported.", , "Nike Women's Fit Reflective Hat"
End Sub

Private Sub cmdInfo4_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Lebron Las Vegas All Star Fitted Cap is made of 100% polyester with a 3D embroidered front logo. Imported.", , "Nike Lebron Vegas All Star Cap"
End Sub

Private Sub cmdInfo5_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Women's Performance Hatphones are a rockin' addition to any runner's wardrobe! Your music player plugs in and pipes in the music. No more headphones bouncing off or heating you up! Therma-FIT® 96% polyester/4% spandex terry. Imported.", , "Nike Women's Perf Hatphones"
End Sub

Private Sub cmdInfo6_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Women's Shiny Mesh Cap is a lightweight, well-ventilated, classic running hat. With an integrated sweat band, mini-mesh back panels for excellent ventilation and 360º reflectivity, this cap is a fast favorite. The comfort bill and one-hand, quick-adjust closure are Nike exclusives. Body/bill: 100% polyester mesh. Sweat band: 100% recycled polyester taffeta. Imported.", , "Nike Women's Shiny Mesh Cap"
End Sub

Private Sub cmdInfo7_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Women's FIT Reflective Running Hat is a choice cap for your run or jog because of all its great features. Built-in terry sweatband controls moisture along with Dri-FIT® materials and mesh panels for ventilation. All-over reflective print provides 360-degree reflectivity, a must-have for nighttime or pre-dawn runs. Quick adjusting closure can be operated with one hand. Body/bill: Dri-FIT® 100% recycled polyester taffeta. Mesh: Dri-FIT® 100% polyester circular knit mesh. Sweatband: Dri-FIT® 90% polyester/10% spandex. Imported.", , "Nike Women's Fit Reflective Hat"
End Sub

Private Sub cmdInfo8_Click()    'allows the user to view the specific information on the item
MsgBox "Jump head first into the Nike Women's Cable Embossed Feather-Light Hat. This runner's hat features exclusive Nike technologies, including a one-hand quick adjustment and comfort bill. The body and bill are made of 100% recycled polyester taffeta. The panels are 100% Dri-FIT® polyester mesh, and the sweatband is 90% polyester/10% spandex terry. It has a reflective silver sandwich bill and dot prints, along with a Swoosh design trademark on the front. Imported.", , "Nike Women's Cable Embossed FTHR Light Hat"
End Sub
