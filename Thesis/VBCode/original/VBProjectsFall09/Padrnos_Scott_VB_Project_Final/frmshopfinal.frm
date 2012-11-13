VERSION 5.00
Begin VB.Form frmshopfinal 
   BackColor       =   &H80000008&
   Caption         =   "Form1"
   ClientHeight    =   13380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20955
   LinkTopic       =   "Form1"
   ScaleHeight     =   13380
   ScaleWidth      =   20955
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtclick 
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Text            =   "Please Click on the item desired"
      Top             =   720
      Width           =   5415
   End
   Begin VB.TextBox txtpride 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      Text            =   "Get Your Johnnie Pride Gear"
      Top             =   120
      Width           =   7215
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to home page"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   15360
      TabIndex        =   4
      Top             =   10560
      Width           =   2175
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   18120
      TabIndex        =   3
      Top             =   10560
      Width           =   2175
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New Transaction"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   18120
      TabIndex        =   2
      Top             =   9360
      Width           =   2175
   End
   Begin VB.CommandButton cmdtotal 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   15360
      TabIndex        =   1
      Top             =   9360
      Width           =   2175
   End
   Begin VB.PictureBox picresults 
      Height          =   8295
      Left            =   15120
      ScaleHeight     =   8235
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   480
      Width           =   5295
   End
   Begin VB.Image Image10 
      Height          =   2835
      Left            =   10680
      Picture         =   "frmshopfinal.frx":0000
      Top             =   5400
      Width           =   3000
   End
   Begin VB.Image Image9 
      Height          =   2850
      Left            =   2160
      Picture         =   "frmshopfinal.frx":1A0C
      Top             =   5280
      Width           =   3000
   End
   Begin VB.Image Image8 
      Height          =   3000
      Left            =   10680
      Picture         =   "frmshopfinal.frx":35B8
      Top             =   9120
      Width           =   3000
   End
   Begin VB.Image Image7 
      Height          =   2895
      Left            =   6600
      Picture         =   "frmshopfinal.frx":4FE2
      Top             =   1320
      Width           =   3000
   End
   Begin VB.Image Image6 
      Height          =   2955
      Left            =   6600
      Picture         =   "frmshopfinal.frx":71A3
      Top             =   9120
      Width           =   3000
   End
   Begin VB.Image Image5 
      Height          =   3045
      Left            =   6600
      Picture         =   "frmshopfinal.frx":9010
      Top             =   5280
      Width           =   3000
   End
   Begin VB.Image Image4 
      Height          =   2895
      Left            =   2160
      Picture         =   "frmshopfinal.frx":AE7A
      Top             =   1320
      Width           =   3000
   End
   Begin VB.Image Image3 
      Height          =   2475
      Left            =   10680
      Picture         =   "frmshopfinal.frx":C2EA
      Top             =   1560
      Width           =   3000
   End
   Begin VB.Image Image2 
      Height          =   2775
      Left            =   2160
      Picture         =   "frmshopfinal.frx":D801
      Top             =   9240
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   7245
      Left            =   840
      Picture         =   "frmshopfinal.frx":E7F1
      Top             =   1320
      Width           =   14235
   End
End
Attribute VB_Name = "frmshopfinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim runningtotal As Single 'declaring all variables

Private Sub cmdnew_Click()
picresults.Cls 'clearing list for when clicked for a second time
runningtotal = 0 'setting the running total to zero
End Sub

Private Sub cmdQuit_Click()
    End 'ending the program
End Sub

Private Sub cmdreturn_Click()
    frmshopfinal.Hide 'hiding the shop form
    frmHome.Show 'returning to the ome form
    'returns to home page
End Sub

Private Sub cmdtotal_Click()
Dim subtotal As Single, tax As Single, total As Single 'declaring all the variables
subtotal = runningtotal 'setting the subtotal
tax = subtotal * 0.07 'setting equation for tax
total = subtotal + tax 'setting the total cost
picresults.Print "____________________________________________________________" 'picresults
picresults.Print "SubTotal"; Tab(45); FormatCurrency(subtotal)
picresults.Print "Tax"; Tab(45); FormatCurrency(tax)
picresults.Print "Total"; Tab(45); FormatCurrency(total)
End Sub
Private Sub Image10_Click()
Dim tshirt3 As Single 'declaring the variables
tshirt3 = 14.99 'setting the price for item
runningtotal = runningtotal + tshirt3
picresults.Print "Screen print grey SJU T-Shirt"; Tab(45); FormatCurrency(tshirt3) 'printing out the item and how much it costs
End Sub

Private Sub Image2_Click()
Dim hat1 As Single
hat1 = 9.99 'setting the price for item
runningtotal = runningtotal + hat1 'adding the item to the runningtotal
picresults.Print "White SJU rat hat"; Tab(45); FormatCurrency(hat1) 'printing out the item and how much it costs
End Sub

Private Sub Image3_Click()
Dim hat2 As Single
hat2 = 27.99 'setting the price for item
runningtotal = runningtotal + hat2 'adding the item to the runningtotal
picresults.Print "Under Armour SJU hat"; Tab(45); FormatCurrency(hat2) 'printing out the item and how much it costs
End Sub

Private Sub Image4_Click()
Dim pants As Single
pants = 29.99 'setting the price for item
runningtotal = runningtotal + pants 'adding the item to the runningtotal
picresults.Print "Screen print SJU sweatpants"; Tab(45); FormatCurrency(pants) 'printing out the item and how much it costs
End Sub

Private Sub Image5_Click()
Dim sweatshirt3 As Single
sweatshirt3 = 49.99 'setting the price for item
runningtotal = runningtotal + sweatshirt3 'adding the item to the runningtotal
picresults.Print "Red two line SJU sweatshirt"; Tab(45); FormatCurrency(sweatshirt3) 'printing out the item and how much it costs
End Sub

Private Sub Image6_Click()
Dim sweatshirt2 As Single
sweatshirt2 = 49.99 'setting the price for item
runningtotal = runningtotal + sweatshirt2 'adding the item to the runningtotal
picresults.Print "Grey SJU 1 sweatshirt"; Tab(45); FormatCurrency(sweatshirt2) 'printing out the item and how much it costs
End Sub

Private Sub Image7_Click()
Dim sweatshirt1 As Single
sweatshirt1 = 35.99 'setting the price for item
runningtotal = runningtotal + sweatshirt1 'adding the item to the runningtotal
picresults.Print "Grey two line SJU sweatshirt"; Tab(45); FormatCurrency(sweatshirt1) 'printing out the item and how much it costs
End Sub

Private Sub Image8_Click()
Dim tshirt1 As Single
tshirt1 = 14.99 'setting the price for item
runningtotal = runningtotal + tshirt1 'adding the item to the runningtotal
picresults.Print "Red SJU T-Shirt"; Tab(45); FormatCurrency(tshirt1) 'printing out the item and how much it costs
End Sub

Private Sub Image9_Click()
Dim tshirt2 As Single
tshirt2 = 23.99 'setting the price for item
runningtotal = runningtotal + tshirt2 'adding the item to the runningtotal
picresults.Print "Sceen print longsleeve Collegeville shirt"; Tab(45); FormatCurrency(tshirt2) 'printing out the item and how much it costs
End Sub

Private Sub txtclick_Change()

End Sub
