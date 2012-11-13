VERSION 5.00
Begin VB.Form frmSweaters 
   Caption         =   "Sweaters/Jackets"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   Picture         =   "frmSweaters.frx":0000
   ScaleHeight     =   8910
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   8640
      TabIndex        =   26
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   495
      Left            =   6720
      TabIndex        =   17
      Top             =   7440
      Width           =   1095
   End
   Begin VB.PictureBox picResults 
      Height          =   3015
      Left            =   6240
      ScaleHeight     =   2955
      ScaleWidth      =   3915
      TabIndex        =   16
      Top             =   4320
      Width           =   3975
   End
   Begin VB.CommandButton cmdInfo8 
      Caption         =   "Info"
      Height          =   495
      Left            =   3840
      TabIndex        =   15
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo7 
      Caption         =   "Info"
      Height          =   495
      Left            =   720
      TabIndex        =   14
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo6 
      Caption         =   "Info"
      Height          =   495
      Left            =   3840
      TabIndex        =   13
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo5 
      Caption         =   "Info"
      Height          =   495
      Left            =   720
      TabIndex        =   12
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo4 
      Caption         =   "Info"
      Height          =   495
      Left            =   8160
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo3 
      Caption         =   "Info"
      Height          =   495
      Left            =   6000
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo2 
      Caption         =   "Info"
      Height          =   495
      Left            =   3720
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo1 
      Caption         =   "Info"
      Height          =   495
      Left            =   840
      TabIndex        =   8
      Top             =   2160
      Width           =   975
   End
   Begin VB.PictureBox Picture8 
      Height          =   1335
      Left            =   600
      Picture         =   "frmSweaters.frx":A056
      ScaleHeight     =   1275
      ScaleWidth      =   1275
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.PictureBox Picture7 
      Height          =   1575
      Left            =   360
      Picture         =   "frmSweaters.frx":A625
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   5520
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   1575
      Left            =   3600
      Picture         =   "frmSweaters.frx":B34C
      ScaleHeight     =   1515
      ScaleWidth      =   1395
      TabIndex        =   5
      Top             =   5520
      Width           =   1455
   End
   Begin VB.PictureBox Picture5 
      Height          =   1575
      Left            =   3600
      Picture         =   "frmSweaters.frx":BBAF
      ScaleHeight     =   1515
      ScaleWidth      =   1395
      TabIndex        =   4
      Top             =   2880
      Width           =   1455
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   600
      Picture         =   "frmSweaters.frx":C8D4
      ScaleHeight     =   1515
      ScaleWidth      =   1395
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.PictureBox Picture3 
      Height          =   1335
      Left            =   8040
      Picture         =   "frmSweaters.frx":D01B
      ScaleHeight     =   1275
      ScaleWidth      =   1275
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Height          =   1335
      Left            =   5760
      Picture         =   "frmSweaters.frx":D5E9
      ScaleHeight     =   1275
      ScaleWidth      =   1275
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   3480
      Picture         =   "frmSweaters.frx":DD1F
      ScaleHeight     =   1275
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "8"
      Height          =   255
      Left            =   3360
      TabIndex        =   25
      Top             =   5520
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "7"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   5520
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "6"
      Height          =   255
      Left            =   3360
      TabIndex        =   23
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "5"
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "4"
      Height          =   255
      Left            =   7800
      TabIndex        =   21
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "3"
      Height          =   255
      Left            =   5520
      TabIndex        =   20
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "2"
      Height          =   255
      Left            =   3120
      TabIndex        =   19
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   480
      Width           =   135
   End
End
Attribute VB_Name = "frmSweaters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmSweaters
'Author: Sean Johnson and Nick Lane
'Date Written: Saturday March 15th, 2007
'Objective of form: this is the women sweaters form and from this form the
'                   user can find out information about the various items available,
'                   buy an item, and view the price of the item along with a running
'                   total of other items.

Dim Sweater(1 To 20) As String
Dim Prices(1 To 20) As Single
Dim ctr As Integer, j As Integer, c As String
Dim found As Boolean, n As Integer, i As Integer

Option Explicit

Private Sub cmdBack_Click()
'this button will hide this form and show the previous form
frmSweaters.Hide
frmWomenApparel.Show
End Sub

'this button will allow the user to purchase an item
Private Sub cmdBuy_Click()
found = False

ctr = 0

'opens the array file to be read into a parallel array
Open App.Path & "\SweaterArray.txt" For Input As #1

'this loop will read the file into a parallel array which will be used when users selects an items
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Sweater(ctr), Prices(ctr)
Loop
Close #1

'allows user to insert the number of the corresponding item he/she wishes to purchase or add to cart.
n = InputBox("Please enter the Number of the corresponding Shoe You wish to purchase")

'this for loop will match the corresponding number above with the corresponding items in the array
For j = 1 To ctr
    If n = j Then
        found = True
        picResults.Print Sweater(j), Tab(25); Prices(j)
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
MsgBox "Jump into fall with Nike's Basic Fleece Hoody. This full-zip fleece hoody is made of 80% cotton (5% organic)/20% polyester which gives a soft and cozy feel in any weather. It also has wire management capabilities, rib detailing and a Nike logo at the left chest. Imported.", , "Nike Women's Basic Fleece Zip Hoody"
End Sub

Private Sub cmdInfo2_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Dri-FIT Light Women's Golf Sweater features a modern, three-quarter length sleeve and pretty ribbing for lightweight layering in style. Nike's Dri-FIT fabric wicks moisture to keep you comfortable through the full eighteen. ", , "Nike Dri-FIT Light Women's Golf Sweater"
End Sub

Private Sub cmdInfo3_Click()    'allows the user to view the specific information on the item
MsgBox "Pre or post match, stay warm in the Nike Drop Shot Women's Cable Sweater, a full-zip jersey knit style with a snap placket and cozy foldover neck. Dri-FIT panels at the underarms and sides help keep you cool under pressure. ", , "Nike Drop Shot Women's Cable Sweater"
End Sub

Private Sub cmdInfo4_Click()    'allows the user to view the specific information on the item
MsgBox "Cut out the chill with the Nike Ohzone Women's Sweater, an athletic-fit deliciously cozy, lightweight layering piece with the look and feel of a heather sweater and the performance-level warmth of Nike Sphere Thermal fabric. ", , "Nike Ohzone Women's Sweater"
End Sub

Private Sub cmdInfo5_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Wool V-Neck Women's Sweater combines traditional wool knit fabric with a sleek modern shape to create a contemporary classic. Features a pretty pointelle design on back. ", , "Nike Wool V-Neck Women's Sweater"
End Sub

Private Sub cmdInfo6_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Down-With-It Women's Jacket features not just a sleek, feminine fit, but also boasts hidden pockets to keep your personal items close at hand. ", , "Nike Down-With-It Women's Jacket"
End Sub

Private Sub cmdInfo7_Click()    'allows the user to view the specific information on the item
MsgBox "Blend comfort and movement with the Nike Women's Unified Knit Jacket. This Dri-FIT® jacket is a distinctly feminine design that features contemporary color block design. Includes zip-front pockets for warming hands or for storage. 91% polyester/9% spandex Dri-FIT® terry. Imported.", , "Nike Women's Unified Knit Jacket"
End Sub

Private Sub cmdInfo8_Click()    'allows the user to view the specific information on the item
MsgBox "Put on your post-game best with the Nike Slacker Women's Jacket. This full-zip, high collar, long-sleeve jacket sports Dri-FIT fabric to deliver high-tech moisture management to keep you cool and dry and feel soft and comfortable against your skin. Features include a relaxed fit and design lines, full front-zip and front full-zip pockets, and contrast piping along collar and arms to give a fresh, modern look with a splash of color. ", , "Nike Slacker Women's Jacket"
End Sub
