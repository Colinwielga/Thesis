VERSION 5.00
Begin VB.Form frmBuildCompact 
   Caption         =   "Build Compact"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   Picture         =   "frmBuildCompact.frx":0000
   ScaleHeight     =   8070
   ScaleWidth      =   8400
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
   Begin VB.CommandButton cmdKey 
      Caption         =   "Keyless Entry"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton cmdHatch 
      Caption         =   "Hatchback"
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
   Begin VB.CommandButton cmdAWD 
      Caption         =   "All-Wheel Drive"
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
Attribute VB_Name = "frmBuildCompact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name is WickedFunCarBuilder
'Form Name is frmBuildCompact
'Author is Dan Parker
'Date written 10/12/09
'The purpose of this form is to let the user build a compact car
Dim itemTotal As Single 'dim global variable

Private Sub cmdClear_Click()

    picResults.Cls 'clears the list of options purchased
    
    cmdLeather.Visible = True 'makes all options available to user
    cmdEngine.Visible = True
    cmdSun.Visible = True
    cmdAWD.Visible = True
    cmdNav.Visible = True
    cmdKey.Visible = True
    cmdHatch.Visible = True
    cmdAuto.Visible = True
End Sub

Private Sub cmdEngine_Click()
    Dim Engine As Single 'dim local variable
    Engine = priceOption(8) 'set engine upgrade cost equal to cost given in array
    itemTotal = itemTotal + Engine 'adds the value of engine upgrade to item total
    picResults.Print "Engine Upgrade"; Tab(45); FormatCurrency(Engine) 'shows the option name and price to user
    
    cmdEngine.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdBack_Click()
    frmBuildCompact.Hide
    frmBegin.Show
End Sub

Private Sub cmdSun_Click()
    Dim Sun As Single
    Sun = priceOption(15)
    itemTotal = itemTotal + Sun 'adds the value of the Tow Package to item total
    picResults.Print "Sunroof"; Tab(45); FormatCurrency(Sun) 'shows the option name and price to user
    
    cmdSun.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdAWD_Click()
    Dim four As Single
    four = priceOption(6)
    itemTotal = itemTotal + four 'adds value of all-wheel drive to item total
    picResults.Print "All-Wheel Drive"; Tab(45); FormatCurrency(four) 'shows the option name and price to user
    
    cmdAWD.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdNav_Click()
    Dim nav As Single
    nav = priceOption(5)
    itemTotal = itemTotal + nav 'adds value of the nav system to item total
    picResults.Print "Navigation System"; Tab(45); FormatCurrency(nav) 'shows the option name and price to user
    
    cmdNav.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdQuit_Click()
MsgBox ("Thanks for using the Wicked Fun Car Builder, " & " " & UserName & "!")
End 'ends program
End Sub

Private Sub cmdKey_Click()
    Dim key As Single
    key = priceOption(17)
    itemTotal = itemTotal + key 'adds value of keyless entry to item total
    picResults.Print "Keyless Entry"; Tab(45); FormatCurrency(key) 'shows the option name and price to user
    
    cmdKey.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdHatch_Click()
    Dim hatch As Single
    hatch = priceOption(13)
    itemTotal = itemTotal + hatch 'adds value of hatchback to item total
    picResults.Print "Hatchback"; Tab(45); FormatCurrency(hatch) 'shows the option name and price to user
    
    cmdHatch.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdAuto_Click()
    Dim auto As Single
    auto = priceOption(1)
    itemTotal = itemTotal + auto 'adds value of auto transmission to item total
    picResults.Print "Automatic Transmission"; Tab(45); FormatCurrency(auto) 'shows the option name and price to user
    
    cmdAuto.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdLeather_Click()
    Dim leather As Single
    leather = priceOption(2)
    itemTotal = itemTotal + leather 'adds value of leather seats to item total
    picResults.Print "Leather Seats"; Tab(45); FormatCurrency(leather) 'shows the option name and price to user
    
    cmdLeather.Visible = False 'hides cmd button to user
End Sub

Private Sub cmdTotal_Click()
    Dim Total As Single, tax As Integer, subTotal As Single
    subTotal = itemTotal + price(5) 'adds price of options to base price of vehicle
    tax = 0.07 * subTotal
    Total = subTotal + tax 'adds tax to subtotal
    
    picResults.Print "   " 'prints blank line
    picResults.Print "   " 'prints blank line
    picResults.Print "Base price of vehicle"; Tab(45); FormatCurrency(price(5)) 'shows base price of car to user as currency
    picResults.Print "Option Total"; Tab(45); FormatCurrency(itemTotal) 'shows option total to user as currency
    picResults.Print "--------------------------------------------------------------------------------------------------"
    picResults.Print "Subtotal"; Tab(45); FormatCurrency(subTotal) 'prints subtotal
    picResults.Print "Tax"; Tab(45); FormatCurrency(tax, 2) 'prints tax
    picResults.Print "--------------------------------------------------------------------------------------------------"
    picResults.Print "Total"; Tab(45); FormatCurrency(Total) 'shows overall total to user as currency
End Sub
