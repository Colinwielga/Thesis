VERSION 5.00
Begin VB.Form frmBuildSedan 
   Caption         =   "Build Sedan"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   Picture         =   "frmBuildSedan.frx":0000
   ScaleHeight     =   7920
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   5055
      Left            =   3600
      ScaleHeight     =   4995
      ScaleWidth      =   4275
      TabIndex        =   12
      Top             =   480
      Width           =   4335
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "Automatic Transmission"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton cmdLeather 
      Caption         =   "Leather Seats"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton cmdClimate 
      Caption         =   "Auto Climate Control"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdLED 
      Caption         =   "LED Headlights"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Navigation System"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton cmdAWD 
      Caption         =   "All-Wheel Drive"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton cmdSun 
      Caption         =   "Sunroof"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton cmdEngine 
      Caption         =   "Engine Upgrade"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   6600
      Width           =   975
   End
End
Attribute VB_Name = "frmBuildSedan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name is WickedFunCarBuilder
'Form Name is frmBuildSedan
'Author is Dan Parker
'Date written 10/12/09
'The purpose of this form is to let the user build a coupe
Dim itemTotal As Single

Private Sub cmdClear_Click()

    picResults.Cls 'clears the list of options purchased
    
    cmdLeather.Visible = True 'makes all options available to user
    cmdEngine.Visible = True
    cmdSun.Visible = True
    cmdAWD.Visible = True
    cmdNav.Visible = True
    cmdLED.Visible = True
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
    frmBuildSedan.Hide
    frmBegin.Show
End Sub

Private Sub cmdSun_Click()
    Dim Sun As Single
    Sun = priceOption(15)
    itemTotal = itemTotal + Sun 'adds the value of a sunroof to item total
    picResults.Print "Sunroof"; Tab(45); FormatCurrency(Sun) 'shows the item name and price to user
    
    cmdSun.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdAWD_Click()
    Dim four As Single
    four = priceOption(6)
    itemTotal = itemTotal + four 'adds value of all-wheel drive to item total
    picResults.Print "All-Wheel Drive"; Tab(45); FormatCurrency(four) 'shows the item name and price to user
    
    cmdAWD.Visible = False 'hides cmd button to user
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

Private Sub cmdLED_Click()
    Dim LED As Single
    LED = priceOption(4)
    itemTotal = itemTotal + LED 'adds value of LED lights to item total
    picResults.Print "LED Headlights"; Tab(45); FormatCurrency(LED) 'shows the item name and price to user
    
    cmdLED.Visible = False 'hides cmd button to user
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
    subTotal = itemTotal + price(2) 'adds price of options to base price of vehicle
    tax = 0.07 * subTotal
    Total = subTotal + tax 'adds tax to subtotal
    
    picResults.Print "   " 'prints blank line
    picResults.Print "   " 'prints blank line
    picResults.Print "Base price of vehicle"; Tab(45); FormatCurrency(price(2)) 'shows base price of car to user as currency
    picResults.Print "Option Total"; Tab(45); FormatCurrency(itemTotal) 'shows option total to user as currency
    picResults.Print "--------------------------------------------------------------------------------------------------"
    picResults.Print "Subtotal"; Tab(45); FormatCurrency(subTotal)
    picResults.Print "Tax"; Tab(45); FormatCurrency(tax)
    picResults.Print "--------------------------------------------------------------------------------------------------"
    picResults.Print "Total"; Tab(45); FormatCurrency(Total) 'shows overall total to user as currency
End Sub

