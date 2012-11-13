VERSION 5.00
Begin VB.Form frmBuildSubcompact 
   Caption         =   "Build Subcompact"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   Picture         =   "frmBuildSubcompact.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   5055
      Left            =   3720
      ScaleHeight     =   4995
      ScaleWidth      =   4275
      TabIndex        =   12
      Top             =   480
      Width           =   4335
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "Automatic Transmission"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton cmdAir 
      Caption         =   "Air Conditioning"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton cmdHatch 
      Caption         =   "Hatchback"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdPower 
      Caption         =   "Power windows and locks"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Navigation System"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton cmdMats 
      Caption         =   "Floor Mats"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton cmdCD 
      Caption         =   "CD Player"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton cmdEngine 
      Caption         =   "Engine Upgrade"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   6600
      Width           =   975
   End
End
Attribute VB_Name = "frmBuildSubcompact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name is WickedFunCarBuilder
'Form Name is frmBuildSubcompact
'Author is Dan Parker
'Date written 10/12/09
'The purpose of this form is to let the user build a subcompact
Dim itemTotal As Single

Private Sub cmdClear_Click()

    picResults.Cls 'clears the list of options purchased
    
    cmdCD.Visible = True 'makes all options available to user
    cmdEngine.Visible = True
    cmdMats.Visible = True
    cmdPower.Visible = True
    cmdNav.Visible = True
    cmdHatch.Visible = True
    cmdAir.Visible = True
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
    frmBuildSubcompact.Hide
    frmBegin.Show
End Sub

Private Sub cmdPower_Click()
    Dim power As Single
    power = priceOption(12)
    itemTotal = itemTotal + power 'adds the value of power windows and locks to item total
    picResults.Print "Power windows and locks"; Tab(45); FormatCurrency(power) 'shows the item name and price to user
    
    cmdPower.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdMats_Click()
    Dim mat As Single
    mat = priceOption(11)
    itemTotal = itemTotal + mat 'adds value of floor mats to item total
    picResults.Print "Floor Mats"; Tab(45); FormatCurrency(mat) 'shows the item name and price to user
    
    cmdMats.Visible = False 'hides cmd button to user
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

Private Sub cmdHatch_Click()
    Dim hatch As Single
    hatch = priceOption(13)
    itemTotal = itemTotal + hatch 'adds value of a hatchback to item total
    picResults.Print "Hatchback"; Tab(45); FormatCurrency(hatch) 'shows the item name and price to user
    
    cmdHatch.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdAir_Click()
    Dim air As Single
    air = priceOption(3)
    itemTotal = itemTotal + air 'adds value of air conditioning to item total
    picResults.Print "Air Conditioning"; Tab(45); FormatCurrency(air) 'shows the item name and price to user
    
    cmdAir.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdAuto_Click()
    Dim auto As Single
    auto = priceOption(1)
    itemTotal = itemTotal + auto 'adds value of auto transmission to item total
    picResults.Print "Automatic Transmission"; Tab(45); FormatCurrency(auto) 'shows the item name and price to user
    
    cmdAuto.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdCD_Click()
    Dim cd As Single
    cd = priceOption(10)
    itemTotal = itemTotal + cd 'adds value of cd player to item total
    picResults.Print "CD Player"; Tab(45); FormatCurrency(cd) 'shows the item name and price to user
    
    cmdCD.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdTotal_Click()
    Dim Total As Single, tax As Integer, subTotal As Single
    subTotal = itemTotal + price(6) 'adds price of options to base price of vehicle
    tax = 0.07 * subTotal
    Total = subTotal + tax 'adds tax to subtotal
    
    picResults.Print "   " 'prints blank line
    picResults.Print "   " 'prints blank line
    picResults.Print "Base price of vehicle"; Tab(45); FormatCurrency(price(6)) 'shows base price of car to user as currency
    picResults.Print "Option Total"; Tab(45); FormatCurrency(itemTotal) 'shows option total to user as currency
    picResults.Print "--------------------------------------------------------------------------------------------------"
    picResults.Print "Subtotal"; Tab(45); FormatCurrency(subTotal)
    picResults.Print "Tax"; Tab(45); FormatCurrency(tax)
    picResults.Print "--------------------------------------------------------------------------------------------------"
    picResults.Print "Total"; Tab(45); FormatCurrency(Total) 'shows overall total to user as currency
End Sub
