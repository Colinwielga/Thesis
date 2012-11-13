VERSION 5.00
Begin VB.Form frmBuildTruck 
   Caption         =   "Build Truck"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   Picture         =   "frmBuildTruck.frx":0000
   ScaleHeight     =   7815
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   5055
      Left            =   3960
      ScaleHeight     =   4995
      ScaleWidth      =   4275
      TabIndex        =   12
      Top             =   600
      Width           =   4335
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "Automatic Transmission"
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdLeather 
      Caption         =   "Leather Seats"
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton cmdClimate 
      Caption         =   "Auto Climate Control"
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton cmdBed 
      Caption         =   "Bed Liner"
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Navigation System"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton cmd4x4 
      Caption         =   "Four-Wheel Drive"
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CommandButton cmdTow 
      Caption         =   "Towing Package"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CommandButton cmdEngine 
      Caption         =   "Engine Upgrade"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   5040
      Width           =   2295
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   7680
      TabIndex        =   2
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   6720
      Width           =   975
   End
End
Attribute VB_Name = "frmBuildTruck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name is WickedFunCarBuilder
'Form Name is frmBuildTruck
'Author is Dan Parker
'Date written 10/12/09
'The purpose of this form is to let the user build a truck
Dim itemTotal As Single

Private Sub cmdClear_Click()

    picResults.Cls 'clears the list of options purchased
    
    cmdLeather.Visible = True 'makes all options available to user
    cmdEngine.Visible = True
    cmdTow.Visible = True
    cmd4x4.Visible = True
    cmdNav.Visible = True
    cmdBed.Visible = True
    cmdClimate.Visible = True
    cmdAuto.Visible = True
End Sub

Private Sub cmdEngine_Click()
    Dim Engine As Single
    Engine = priceOption(8)
    itemTotal = itemTotal + Engine 'adds the value of engine upgrade to item total
    picResults.Print "Engine Upgrade"; Tab(45); FormatCurrency(Engine) 'shows the item name and price to user
    
    cmdEngine.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdBack_Click()
    frmBuildTruck.Hide
    frmBegin.Show
End Sub

Private Sub cmdTow_Click()
    Dim Tow As Single
    Tow = priceOption(7)
    itemTotal = itemTotal + Tow 'adds the value of the Tow Package to item total
    picResults.Print "Tow Package"; Tab(45); FormatCurrency(Tow) 'shows the item name and price to user
    
    cmdTow.Visible = False 'hides cmd button to user
End Sub

Private Sub cmd4x4_Click()
    Dim four As Single
    four = priceOption(6)
    itemTotal = itemTotal + four 'adds value of 4x4 to item total
    picResults.Print "4 x 4"; Tab(45); FormatCurrency(four) 'shows the item name and price to user
    
    cmd4x4.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdNav_Click()
    Dim nav As Single
    nav = priceOption(5)
    itemTotal = itemTotal + nav 'adds value of the nav system to item total
    picResults.Print "Navigation System"; Tab(45); FormatCurrency(nav) 'shows the item name and price to user
    
    cmdNav.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdQuit_Click()
MsgBox ("Thanks for using the Wicked Fun Car Builder, " & " " & UserName & "!")
End
End Sub

Private Sub cmdBed_Click()
    Dim bed As Single
    bed = priceOption(9)
    itemTotal = itemTotal + bed 'adds value of bed liner to item total
    picResults.Print "Bed Liner"; Tab(45); FormatCurrency(bed) 'shows the item name and price to user
    
    cmdBed.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdClimate_Click()
    Dim climate As Single
    climate = priceOption(3)
    itemTotal = itemTotal + climate 'adds value of Auto Climate Control to item total
    picResults.Print "Automatic Climate Control"; Tab(45); FormatCurrency(climate) 'shows the item name and price to user
    
    cmdClimate.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdAuto_Click()
    Dim auto As Single
    auto = priceOption(1)
    itemTotal = itemTotal + auto 'adds value of auto transmission to item total
    picResults.Print "Automatic Transmission"; Tab(45); FormatCurrency(auto) 'shows the item name and price to user
    
    cmdAuto.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdLeather_Click()
    Dim leather As Single
    leather = priceOption(2)
    itemTotal = itemTotal + leather 'adds value of leather seats to item total
    picResults.Print "Leather Seats"; Tab(45); FormatCurrency(leather) 'shows the item name and price to user
    
    cmdLeather.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdTotal_Click()
    Dim Total As Single, tax As Integer, subTotal As Single
    subTotal = itemTotal + price(3) 'adds price of options to base price of vehicle
    tax = 0.07 * subTotal
    Total = subTotal + tax 'adds tax to subtotal

    picResults.Print "   " 'prints blank line
    picResults.Print "   " 'prints blank line
    picResults.Print "Base price of vehicle"; Tab(45); FormatCurrency(price(3)) 'shows base price of car to user as currency
    picResults.Print "Option Total"; Tab(45); FormatCurrency(itemTotal) 'shows option total to user as currency
    picResults.Print "--------------------------------------------------------------------------------------------------"
    picResults.Print "Subtotal"; Tab(45); FormatCurrency(subTotal)
    picResults.Print "Tax"; Tab(45); FormatCurrency(tax)
    picResults.Print "--------------------------------------------------------------------------------------------------"
    picResults.Print "Total"; Tab(45); FormatCurrency(Total) 'shows overall total to user as currency
End Sub
