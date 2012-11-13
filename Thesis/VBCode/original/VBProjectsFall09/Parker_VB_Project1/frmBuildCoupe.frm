VERSION 5.00
Begin VB.Form frmBuildCoupe 
   Caption         =   "Build Coupe"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   Picture         =   "frmBuildCoupe.frx":0000
   ScaleHeight     =   7875
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   5055
      Left            =   3720
      ScaleHeight     =   4995
      ScaleWidth      =   4275
      TabIndex        =   12
      Top             =   360
      Width           =   4335
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "Automatic Transmission"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton cmdLeather 
      Caption         =   "Leather Seats"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton cmdClimate 
      Caption         =   "Auto Climate Control"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton cmdLED 
      Caption         =   "LED Headlights"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Navigation System"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton cmdSport 
      Caption         =   "Sport Package"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton cmdSun 
      Caption         =   "Sunroof"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton cmdEngine 
      Caption         =   "Engine Upgrade"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4800
      Width           =   2295
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   6480
      Width           =   975
   End
End
Attribute VB_Name = "frmBuildCoupe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name is WickedFunCarBuilder
'Form Name is frmBuildCoupe
'Author is Dan Parker
'Date written 10/12/09
'The purpose of this form is to let the user build a coupe
Dim itemTotal As Single

Private Sub cmdClear_Click()

    picResults.Cls 'clears the list of options purchased
    
    cmdLeather.Visible = True 'makes all options available to user
    cmdEngine.Visible = True
    cmdSun.Visible = True
    cmdSport.Visible = True
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
    frmBuildCoupe.Hide
    frmBegin.Show
End Sub

Private Sub cmdSun_Click()
    Dim Sun As Single
    Sun = priceOption(15)
    itemTotal = itemTotal + Sun 'adds the value of a sunroof to item total
    picResults.Print "Sunroof"; Tab(45); FormatCurrency(Sun) 'shows the item name and price to user
    
    cmdSun.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdSport_Click()
    Dim sport As Single
    sport = priceOption(16)
    itemTotal = itemTotal + sport 'adds value of a Sport Package to item total
    picResults.Print "All-Wheel Drive"; Tab(45); FormatCurrency(sport) 'shows the item name and price to user
    
    cmdSport.Visible = False 'hides cmd button to user
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
End 'ends program
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
    subTotal = itemTotal + price(4) 'adds price of options to base price of vehicle
    tax = 0.07 * subTotal
    Total = subTotal + tax 'adds tax to subtotal
    
    picResults.Print "   " 'prints blank line
    picResults.Print "   " 'prints blank line
    picResults.Print "Base price of vehicle"; Tab(45); FormatCurrency(price(4)) 'shows base price of car to user as currency
    picResults.Print "Option Total"; Tab(45); FormatCurrency(itemTotal) 'shows option total to user as currency
    picResults.Print "--------------------------------------------------------------------------------------------------"
    picResults.Print "Subtotal"; Tab(45); FormatCurrency(subTotal)
    picResults.Print "Tax"; Tab(45); FormatCurrency(tax)
    picResults.Print "--------------------------------------------------------------------------------------------------"
    picResults.Print "Total"; Tab(45); FormatCurrency(Total) 'shows overall total to user as currency
End Sub


