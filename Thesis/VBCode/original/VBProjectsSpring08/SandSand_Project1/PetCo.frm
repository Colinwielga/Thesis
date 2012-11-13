VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H00FF8080&
   Caption         =   "frmMainMenu"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H008080FF&
      Caption         =   "Empty Shopping Cart"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4560
      Width           =   2895
   End
   Begin VB.CommandButton cmdShoppingCart 
      BackColor       =   &H008080FF&
      Caption         =   "Show Items"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdPurchase 
      BackColor       =   &H008080FF&
      Caption         =   "Purchase"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CommandButton cmdFishSup 
      BackColor       =   &H00FF0000&
      Caption         =   "Fish Supplies"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7920
      Width           =   2295
   End
   Begin VB.CommandButton cmdCatSup 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cat Supplies"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7920
      Width           =   2295
   End
   Begin VB.CommandButton cmdDogSup 
      BackColor       =   &H00000080&
      Caption         =   "Dog Supplies"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7920
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      Height          =   2775
      Left            =   11760
      Picture         =   "PetCo.frx":0000
      ScaleHeight     =   2715
      ScaleWidth      =   2355
      TabIndex        =   7
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton cmdFish 
      BackColor       =   &H00FF0000&
      Caption         =   "Fish"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4155
      ScaleWidth      =   5235
      TabIndex        =   5
      Top             =   840
      Width           =   5295
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12600
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9120
      Width           =   2295
   End
   Begin VB.CommandButton cmdCats 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cats"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   1
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CommandButton cmdDogs 
      BackColor       =   &H00000080&
      Caption         =   "Dogs"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   0
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Begin Shopping Here >>"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   21.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   15
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label LabPurchaseOrder 
      BackColor       =   &H00FF8080&
      Caption         =   "Here is what is in your shopping cart: "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label labNames 
      BackColor       =   &H00FF8080&
      Caption         =   "By: Scott Sand and Kate Sand"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   4
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label LabTitle 
      BackColor       =   &H00FF8080&
      Caption         =   "    Welcome to Sand's Pet Store"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   1815
      Left            =   8040
      TabIndex        =   3
      Top             =   360
      Width           =   5655
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sand's Pet Store
'Name of Form: frmMainMenu
'Author: Scott Sand and Kate Sand
'Date Written: March 7, 2008
'Objective: Is to welcome custumers to Sand's Pet Store. It also is where they check out when finished shopping.
'Other Comments: The custumers can click on dogs, cats, or fish to
'                to see the types of animals Sand's Pet Store has.
'                If they don't want to purchas a pet, they can skip
'                to house or toys for their pet.

Option Explicit

Private Sub cmdCats_Click(Index As Integer)
' Directs customer to the cat selection form
frmTypesCats.Show
frmMainMenu.Hide
End Sub

Private Sub cmdCatSup_Click()
'Directs the customer to the litter boxes and other supplies
frmLitterBox.Show
frmMainMenu.Hide
End Sub

Private Sub cmdClear_Click()
'Clears all items in the shopping cart and the totals in the module
PicResults.Cls
HabitatCost = 0
FoodCost = 0
PetCost = 0
AccesoriesCost = 0
WeeksSupply = 0
End Sub

Private Sub cmdDogs_Click(Index As Integer)
'Directs customers to the dog selection form
frmTypesDogs.Show
frmMainMenu.Hide
End Sub

Private Sub cmdDogSup_Click()
'Directs the customer to the dog kennels and other supplies
frmDogKennels.Show
frmMainMenu.Hide
End Sub

Private Sub cmdFish_Click()
' Directs customers to the fish selection form
frmTypesFish.Show
frmMainMenu.Hide
End Sub

Private Sub cmdFishSup_Click()
'Directs customers to the fish tanks and other supplies
frmFishTanks.Show
frmMainMenu.Hide
End Sub

Private Sub cmdPurchase_Click()
'The final step of the process: the customer selects to purchase then the project closes
GrandTotal = (NetTotal * 0.065) + NetTotal
MsgBox ("Your Grand total with 6.5% tax is " & FormatCurrency(GrandTotal) & ", and your order will be shipped to you in four business days. Thank You for shopping at Sand's Pet Store.")
End
End Sub

Private Sub cmdQuit_Click()
'Ends program
End
End Sub

Private Sub Command1_Click()
inputdog = txtDog.Text
PicResults.Print inputdog
End Sub

Private Sub cmdShoppingCart_Click()
'Prints list of all items in the shopping cart so that the customer can review them before purchasing
PicResults.Cls
NetTotal = PetCost + HabitatCost + AccesoriesCost + FoodCost
PicResults.Print "Type of Cost"; Tab(35); "Cost"
PicResults.Print "*****************************************************"
PicResults.Print "Pet "; Tab(35); FormatCurrency(PetCost)
PicResults.Print "Habitat "; Tab(35); FormatCurrency(HabitatCost)
PicResults.Print "Accesories "; Tab(35); FormatCurrency(AccesoriesCost)
PicResults.Print "Food"; Tab(35); FormatCurrency(FoodCost)
PicResults.Print "                    "
PicResults.Print "-----------------------------------------------------------"
PicResults.Print " Your pre-tax total is "; FormatCurrency(NetTotal)

End Sub
