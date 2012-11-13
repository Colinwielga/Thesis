VERSION 5.00
Begin VB.Form frmFishFood 
   BackColor       =   &H00FF0000&
   Caption         =   "Fish Food"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdNo 
      Caption         =   "No Thank You!"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8040
      TabIndex        =   9
      Top             =   9000
      Width           =   2415
   End
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11040
      TabIndex        =   7
      Top             =   9120
      Width           =   3975
   End
   Begin VB.CommandButton cmdTetra 
      BackColor       =   &H0000C000&
      Caption         =   "Tetra"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   3975
   End
   Begin VB.CommandButton cmdHikari 
      BackColor       =   &H0000C000&
      Caption         =   "Hikari"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   3975
   End
   Begin VB.CommandButton cmdONutrition 
      BackColor       =   &H0000C000&
      Caption         =   "Ocean Nutrition"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF0000&
      Caption         =   "Select a type of Fish Food you would like to purchase.   "
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1095
      Index           =   1
      Left            =   600
      TabIndex        =   8
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Caption         =   "Ocean Nutrition is a very healthy food for valuable fish. A weeks supply costs $8.00."
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   11.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1815
      Index           =   0
      Left            =   1080
      TabIndex        =   6
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   "Tetra is a basic everyday fish food. A weeks supply costs $5.00."
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   11.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1335
      Left            =   10800
      TabIndex        =   5
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      Caption         =   "Hikari is an organic fish food that is very healthy. A weeks supply costs $7.00."
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   11.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1695
      Left            =   5880
      TabIndex        =   4
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Fish Food"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   48
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   2415
      Left            =   1080
      TabIndex        =   3
      Top             =   7800
      Width           =   6255
   End
End
Attribute VB_Name = "frmFishFood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sand's Pet Store
'Name of Form: frmFishFood
'Author: Scott Sand and Kate Sand
'Date Written: March 10, 2008
'Objective: This is where people can choose which type and how much food they would like for their fish.
'Other Comments:

Option Explicit

Private Sub cmdHikari_Click()
'The customer chooses a type of food and how much then a message box calculates the total cost
WeeksSupply = InputBox("How many weeks supply would you like to purchae")
FoodCost = FoodCost + WeeksSupply * 7
MsgBox ("You purchased enough food for " & WeeksSupply & " weeks and the cost is " & FormatCurrency(FoodCost))
frmMainMenu.Show
frmFishFood.Hide
End Sub

Private Sub cmdMainMenu_Click()
'The customer is directed back to the main menu
frmFishFood.Hide
frmMainMenu.Show
End Sub

Private Sub cmdNo_Click()
'The customer can choose not to buy fish food
frmFishFood.Hide
frmMainMenu.Show
End Sub

Private Sub cmdONutrition_Click()
'The customer chooses a type of food and how much then a message box calculates the total cost
WeeksSupply = InputBox("How many weeks supply would you like to purchae")
FoodCost = FoodCost + WeeksSupply * 8
MsgBox ("You purchased enough food for " & WeeksSupply & " weeks and the cost is " & FormatCurrency(FoodCost))
frmMainMenu.Show
frmFishFood.Hide
End Sub

Private Sub cmdTetra_Click()
'The customer chooses a type of food and how much then a message box calculates the total cost
WeeksSupply = InputBox("How many weeks supply would you like to purchae")
FoodCost = FoodCost + WeeksSupply * 5
MsgBox ("You purchased enough food for " & WeeksSupply & " weeks and the cost is " & FormatCurrency(FoodCost))
frmMainMenu.Show
frmFishFood.Hide
End Sub
