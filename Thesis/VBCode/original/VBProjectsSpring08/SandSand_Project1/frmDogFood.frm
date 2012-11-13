VERSION 5.00
Begin VB.Form frmDogFood 
   BackColor       =   &H00000080&
   Caption         =   "Dog Food"
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
      Left            =   7080
      TabIndex        =   11
      Top             =   9120
      Width           =   2295
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1009
      Left            =   10320
      TabIndex        =   4
      Top             =   9240
      Width           =   4575
   End
   Begin VB.CommandButton cmdPedigree 
      BackColor       =   &H0080FFFF&
      Caption         =   "Pedigree"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton cmdEukanuba 
      BackColor       =   &H0080FFFF&
      Caption         =   "Eukanuba"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton cmdIams 
      BackColor       =   &H0080FFFF&
      Caption         =   "Iams"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   3375
   End
   Begin VB.CommandButton cmdNutroNat 
      BackColor       =   &H0080FFFF&
      Caption         =   "Nutro Natural"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000080&
      Caption         =   "Select the type of dog food you would like to purchase.  "
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1815
      Index           =   1
      Left            =   1560
      TabIndex        =   10
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000080&
      Caption         =   "Pedigree is a well known dog food that will provide the essential dietary needs to your dog. A weeks supply costs $10.00."
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1695
      Left            =   11640
      TabIndex        =   9
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000080&
      Caption         =   $"frmDogFood.frx":0000
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1695
      Index           =   0
      Left            =   7920
      TabIndex        =   8
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000080&
      Caption         =   "Iams is a healthy dog food that focuses on the longevity of your dog. A weeks supply costs $12.00."
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1815
      Left            =   4080
      TabIndex        =   7
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000080&
      Caption         =   "Nutro Natural is the healthiest choice of Dog food for you dog. A weeks supply costs $15.00."
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1815
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Caption         =   "Dog Food"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   48
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1215
      Left            =   840
      TabIndex        =   5
      Top             =   8280
      Width           =   4455
   End
End
Attribute VB_Name = "frmDogFood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sand's Pet Store
'Name of Form: frmDogFood
'Author: Scott Sand and Kate Sand
'Date Written: March 12, 2008
'Objective: Allows customers to choose which type and how much dog food they would like to purchase.
'Other Comments:

Option Explicit

Private Sub cmdBack_Click()
'This sends the customer back to the main menu
frmDogFood.Hide
frmMainMenu.Show
End Sub

Private Sub cmdEukanuba_Click()
'A type of food is selected and how much then a total is calculated
WeeksSupply = InputBox("How many weeks supply would you like to purchae")
FoodCost = FoodCost + WeeksSupply * 12
MsgBox ("You purchased enough food for " & WeeksSupply & " weeks and the cost is " & FormatCurrency(FoodCost))
frmMainMenu.Show
frmDogFood.Hide
End Sub

Private Sub cmdIams_Click()
'A type of food is selected and how much then a total is calculated
WeeksSupply = InputBox("How many weeks supply would you like to purchae")
FoodCost = FoodCost + WeeksSupply * 12
MsgBox ("You purchased enough food for " & WeeksSupply & " weeks and the cost is " & FormatCurrency(FoodCost))
frmMainMenu.Show
frmDogFood.Hide
End Sub

Private Sub cmdNo_Click()
'The customer can choose not to buy dog food
frmDogFood.Hide
frmMainMenu.Show
End Sub

Private Sub cmdNutroNat_Click()
'A type of food is selected and how much then a total is calculated
WeeksSupply = InputBox("How many weeks supply would you like to purchae")
FoodCost = FoodCost + WeeksSupply * 15
MsgBox ("You purchased enough food for " & WeeksSupply & " weeks and the cost is " & FormatCurrency(FoodCost))
frmMainMenu.Show
frmDogFood.Hide
End Sub

Private Sub cmdPedigree_Click()
'A type of food is selected and how much then a total is calculated
WeeksSupply = InputBox("How many weeks supply would you like to purchae")
FoodCost = FoodCost + WeeksSupply * 10
MsgBox ("You purchased enough food for " & WeeksSupply & " weeks and the cost is " & FormatCurrency(FoodCost))
frmMainMenu.Show
frmDogFood.Hide
End Sub
