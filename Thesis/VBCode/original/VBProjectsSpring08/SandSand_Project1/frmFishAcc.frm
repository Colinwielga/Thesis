VERSION 5.00
Begin VB.Form frmFishAcc 
   BackColor       =   &H00FF0000&
   Caption         =   "Fish Tank Accessories"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdMove 
      BackColor       =   &H0000C000&
      Caption         =   "Move to Fish Food"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6840
      Width           =   4335
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
      Left            =   10920
      TabIndex        =   5
      Top             =   8880
      Width           =   3975
   End
   Begin VB.CommandButton cmdNoThanks 
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
      Height          =   1095
      Left            =   8400
      TabIndex        =   4
      Top             =   8880
      Width           =   1935
   End
   Begin VB.CommandButton cmdRocks 
      BackColor       =   &H0000C000&
      Caption         =   "Rocks"
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
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   3855
   End
   Begin VB.CommandButton cmdPShip 
      BackColor       =   &H0000C000&
      Caption         =   "Pirate Ship"
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   3855
   End
   Begin VB.CommandButton cmdPlants 
      BackColor       =   &H0000C000&
      Caption         =   "Plants"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label labInstructions 
      BackColor       =   &H00FF0000&
      Caption         =   "Click on the buttons to select what accessories you would like for your fish.  "
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
      Height          =   735
      Left            =   480
      TabIndex        =   6
      Top             =   600
      Width           =   5535
   End
   Begin VB.Label lblFishAcc 
      BackColor       =   &H00FF0000&
      Caption         =   "Fish Tank Accessories"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   2535
      Left            =   1440
      TabIndex        =   3
      Top             =   7080
      Width           =   4815
   End
End
Attribute VB_Name = "frmFishAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sand's Pet Store
'Name of Form: frmFishAcc
'Author: Scott Sand and Kate Sand
'Date Written: March 12, 2008
'Objective: This is where people can find accesories for their fish tanks.
'Other Comments:

Option Explicit

Private Sub cmdMainMenu_Click()
'The customer is directed back to the main menu
frmFishAcc.Hide
frmMainMenu.Show
End Sub

Private Sub cmdMove_Click()
'The customer is finished purchasing fish accesories and wishes to move to the fish food form
frmFishFood.Show
frmFishAcc.Hide
End Sub

Private Sub cmdNoThanks_Click()
'The customer can choose not to buy fish tank accesories
frmFishAcc.Hide
frmFishFood.Show
End Sub

Private Sub cmdPlants_Click()
'Thye customer chooses to buy plants
MsgBox ("You have chosen to purchase Plants for a fish tank that cost $8")
AccesoriesCost = AccesoriesCost + 8
End Sub

Private Sub cmdPShip_Click()
'The  customer chooses to purchase a pirate ship
MsgBox ("You have chosen to purchase a Pirate Ship that costs $9")
AccesoriesCost = AccesoriesCost + 9
End Sub

Private Sub cmdRocks_Click()
'The customer chooses to purchase rocks for the fish tank
MsgBox ("You have chosen to purchase Rocks for a fish tank that cost $7")
AccesoriesCost = AccesoriesCost + 7
End Sub
