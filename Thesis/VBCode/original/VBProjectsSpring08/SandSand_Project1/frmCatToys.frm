VERSION 5.00
Begin VB.Form frmCatToys 
   BackColor       =   &H0080C0FF&
   Caption         =   "Cat Toys"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdMove 
      BackColor       =   &H00FF8080&
      Caption         =   "Move to Cat Food"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7680
      Width           =   4095
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
      Left            =   10680
      TabIndex        =   5
      Top             =   9240
      Width           =   3975
   End
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
      Left            =   8280
      TabIndex        =   4
      Top             =   9240
      Width           =   1935
   End
   Begin VB.CommandButton cmdMice 
      BackColor       =   &H00FF8080&
      Caption         =   "Toy Mice"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   3855
   End
   Begin VB.CommandButton cmdClimbing 
      BackColor       =   &H00FF8080&
      Caption         =   "Climbing Equipment"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   3975
   End
   Begin VB.CommandButton cmdYarn 
      BackColor       =   &H00FF8080&
      Caption         =   "Ball of Yarn"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   4095
   End
   Begin VB.Label labInstructions 
      BackColor       =   &H0080C0FF&
      Caption         =   "Click on the buttons to select what toys you would like for your cat (you can choose more than one).  "
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
      Height          =   1215
      Left            =   1080
      TabIndex        =   7
      Top             =   720
      Width           =   6615
   End
   Begin VB.Label lblCatToys 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cat Toys"
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
      Height          =   1455
      Left            =   600
      TabIndex        =   3
      Top             =   7800
      Width           =   4575
   End
End
Attribute VB_Name = "frmCatToys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sand's Pet Store
'Name of Form: frmCatToys
'Author: Scott Sand and Kate Sand
'Date Written: March 12, 2008
'Objective: This is where customers can choose toys for their cats.
'Other Comments:

Option Explicit

Private Sub cmdBack_Click()
'sends the customer back to the main menu
frmCatToys.Hide
frmMainMenu.Show
End Sub

Private Sub cmdClimbing_Click()
'customer selects to buy this item
MsgBox ("You have chosen to purchase Climbing Equipment that costs $65")
AccesoriesCost = AccesoriesCost + 65
End Sub

Private Sub cmdMice_Click()
'customer selects to buy this item
MsgBox ("You have chosen to purchase a Toy Mice that costs $4")
AccesoriesCost = AccesoriesCost + 4
End Sub

Private Sub cmdMove_Click()
'customer is done choosing toys and can move to the cat food
frmCatFood.Show
frmCatToys.Hide
End Sub

Private Sub cmdNo_Click()
'The customer can choose not to buy food
frmCatToys.Hide
frmCatFood.Show
End Sub

Private Sub cmdYarn_Click()
'customer selects to buy this item
MsgBox ("You have chosen to purchase a Ball of Yarn that costs $2")
AccesoriesCost = AccesoriesCost + 2
End Sub
