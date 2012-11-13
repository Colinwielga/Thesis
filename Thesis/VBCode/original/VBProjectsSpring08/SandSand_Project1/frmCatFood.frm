VERSION 5.00
Begin VB.Form frmCatFood 
   BackColor       =   &H0080C0FF&
   Caption         =   "Cat Food"
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
      Left            =   7680
      TabIndex        =   11
      Top             =   9240
      Width           =   2175
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
      Height          =   975
      Index           =   1
      Left            =   10800
      TabIndex        =   9
      Top             =   9240
      Width           =   3975
   End
   Begin VB.CommandButton cmdFriskies 
      BackColor       =   &H00FF8080&
      Caption         =   "Friskies"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   3135
   End
   Begin VB.CommandButton cmdPOne 
      BackColor       =   &H00FF8080&
      Caption         =   "Purina One"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   3135
   End
   Begin VB.CommandButton cmdIams 
      BackColor       =   &H00FF8080&
      Caption         =   "Iams"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   3015
   End
   Begin VB.CommandButton cmdSDiet 
      BackColor       =   &H00FF8080&
      Caption         =   "Science Diet"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Select a type of cat food to purchase."
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   735
      Index           =   1
      Left            =   600
      TabIndex        =   10
      Top             =   480
      Width           =   6375
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Friskies is a basic cat food that offers all of the dietary needs for your cat. A weeks supply costs $10.00."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   2055
      Index           =   0
      Left            =   11640
      TabIndex        =   8
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Purina One is a healthy food deigned for active cats. A weeks supply costs $12.00."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   2055
      Left            =   7800
      TabIndex        =   7
      Top             =   4080
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Iams is a healthy cat food that is formulated for the longevity of your cats life. A weeks supply costs $12.00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   2055
      Left            =   4320
      TabIndex        =   6
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Science Diet is the Healthiest choice that Petco offers. A weeks supply costs $15.00."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1935
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   4080
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cat Food"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1815
      Left            =   1080
      TabIndex        =   4
      Top             =   8040
      Width           =   4575
   End
End
Attribute VB_Name = "frmCatFood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sand's Pet Store
'Name of Form: frmCatFood
'Author: Scott Sand and Kate Sand
'Date Written: March 12, 2008
'Objective: Allows customer to pick which type and how much cat food they would like to purchase.
'Other Comments:

Option Explicit

Private Sub cmdBack_Click(Index As Integer)

'Directs customer back to the main menu
frmCatFood.Hide
frmMainMenu.Show
End Sub

Private Sub cmdFriskies_Click()
'the customer selects this brand of food and how many weeks supply they want then a total is calculated
WeeksSupply = InputBox("How many weeks supply would you like to purchae")
FoodCost = FoodCost + WeeksSupply * 10
MsgBox ("You purchased enough food for " & WeeksSupply & " weeks and the cost is " & FormatCurrency(FoodCost))
frmMainMenu.Show
frmCatFood.Hide
End Sub

Private Sub cmdIams_Click()
'the customer selects this brand of food and how many weeks supply they want then a total is calculated
WeeksSupply = InputBox("How many weeks supply would you like to purchae")
FoodCost = FoodCost + WeeksSupply * 12
MsgBox ("You purchased enough food for " & WeeksSupply & " weeks and the cost is " & FormatCurrency(FoodCost))
frmMainMenu.Show
frmCatFood.Hide
End Sub

Private Sub cmdNo_Click()
'the customer can choose not to purchase food
frmCatFood.Hide
frmMainMenu.Show
End Sub

Private Sub cmdPOne_Click()
'the customer selects this brand of food and how many weeks supply they want then a total is calculated
WeeksSupply = InputBox("How many weeks supply would you like to purchae")
FoodCost = FoodCost + WeeksSupply * 12
MsgBox ("You purchased enough food for " & WeeksSupply & " weeks and the cost is " & FormatCurrency(FoodCost))
frmMainMenu.Show
frmCatFood.Hide
End Sub

Private Sub cmdSDiet_Click()
'the customer selects this brand of food and how many weeks supply they want then a total is calculated
WeeksSupply = InputBox("How many weeks supply would you like to purchae")
FoodCost = FoodCost + WeeksSupply * 15
MsgBox ("You purchased enough food for " & WeeksSupply & " weeks and the cost is " & FormatCurrency(FoodCost))
frmMainMenu.Show
frmCatFood.Hide
End Sub
