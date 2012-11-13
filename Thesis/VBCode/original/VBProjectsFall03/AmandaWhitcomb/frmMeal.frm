VERSION 5.00
Begin VB.Form frmMeal 
   BackColor       =   &H00404080&
   Caption         =   "Meal"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   975
      Left            =   2160
      TabIndex        =   10
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton cmdChicken 
      Caption         =   "Chicken"
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdSeafood 
      Caption         =   "Sea Food"
      Height          =   975
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdDessert 
      Caption         =   "Dessert"
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   975
      Left            =   8160
      TabIndex        =   6
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Calculate your meal cost"
      Height          =   975
      Left            =   4200
      TabIndex        =   5
      Top             =   5280
      Width           =   1815
   End
   Begin VB.PictureBox mealresults 
      Height          =   3255
      Left            =   3360
      ScaleHeight     =   3195
      ScaleWidth      =   5355
      TabIndex        =   4
      Top             =   600
      Width           =   5415
   End
   Begin VB.CommandButton cmdBeverage 
      Caption         =   "Beverage"
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdSteak 
      Caption         =   "Steak"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdAppetizer 
      Caption         =   "Appetizer"
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   9000
      Picture         =   "frmMeal.frx":0000
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "                                             Choose your meal         "
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmMeal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim total As Double
Dim appetizer As Integer
Dim chicken As Integer
Dim steak As Integer
Dim seafood As Integer
Dim beverage As Integer
Dim dessert As Integer

Private Sub cmdAppetizer_Click()
appetizer = 7
total = total + appetizer
mealresults.Print "appetizer"; Tab; FormatCurrency(appetizer)
End Sub

Private Sub cmdBeverage_Click()
beverage = 4
total = total + beverage
mealresults.Print ; "beverage"; Tab; FormatCurrency(beverage)
End Sub

Private Sub cmdChicken_Click()
chicken = 10
total = total + chicken
mealresults.Print "chicken"; Tab; FormatCurrency(chicken)
End Sub

Private Sub cmdClear_Click()
total = 0
mealresults.Cls
End Sub

Private Sub cmdCompute_Click()
mealresults.Print "------------------------------------------------"
mealresults.Print "subtotal"; Tab; FormatCurrency(total)
mealresults.Print "tip"; Tab; FormatCurrency(total * 0.2)
mealresults.Print "total"; Tab; FormatCurrency(total + (total * 0.2))
totalmeal = total
End Sub

Private Sub cmdContinue_Click()
mealresults.Cls
total = 0

'Hide the Meal selection screen and show
'the Calculate selection screen for the users next input.
frmMeal.Hide
frmCalculate.Show

End Sub

Private Sub cmdDessert_Click()
dessert = 5
total = total + dessert
mealresults.Print "dessert"; Tab; FormatCurrency(dessert)
End Sub

Private Sub cmdSeafood_Click()
seafood = 25
total = total + seafood
mealresults.Print "sea food"; Tab; FormatCurrency(seafood)
End Sub

Private Sub cmdSteak_Click()
steak = 20
total = total + steak
mealresults.Print "steak"; Tab; FormatCurrency(steak)
End Sub
