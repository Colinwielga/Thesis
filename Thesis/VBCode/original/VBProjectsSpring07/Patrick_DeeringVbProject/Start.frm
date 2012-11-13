VERSION 5.00
Begin VB.Form frmHome 
   BackColor       =   &H00000000&
   Caption         =   "Home"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheckout 
      Caption         =   "Check Out"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   9
      Top             =   5160
      Width           =   2655
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9120
      TabIndex        =   8
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton cmdFrozenFoods 
      Caption         =   "Frozen Foods"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   7
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CommandButton cmdToiletries 
      Caption         =   "Toiletries"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   6
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton cmdDeli 
      Caption         =   "Deli"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   5
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton cmdBakery 
      Caption         =   "Bakery"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CommandButton cmdFruit 
      Caption         =   "Fruit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton cmdVegetables 
      Caption         =   "Vegetables"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label lblWarning 
      BackColor       =   &H00000000&
      Caption         =   "*All fields must be filled"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   10
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      X1              =   4560
      X2              =   4560
      Y1              =   2160
      Y2              =   4920
   End
   Begin VB.Label lblWhere 
      BackColor       =   &H00000000&
      Caption         =   "Where would you like to start shopping today?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   9360
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblHome 
      BackColor       =   &H00000000&
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBakery_Click()
frmHome.Hide
frmBakery.Show
'direct user to Bakery form
End Sub

Private Sub cmdCheckout_Click()
frmHome.Hide
frmCheckOut.Show
'direct user to CheckOut form
End Sub

Private Sub cmdDeli_Click()
frmHome.Hide
frmDeli.Show
'direct user to Deli form
End Sub

Private Sub cmdFrozenFoods_Click()
frmHome.Hide
frmFrozenFoods.Show
'direct user to FrozenFoods form
End Sub

Private Sub cmdFruit_Click()
frmHome.Hide
frmFruit.Show
'direct user to Fruit form
End Sub

Private Sub CmdQuit_Click()
End
End Sub

Private Sub cmdToiletries_Click()
frmHome.Hide
frmToiletries.Show
'direct user to Toiletries form
End Sub

Private Sub cmdVegetables_Click()
frmHome.Hide
frmVegetables.Show
'direct user to Vegetables form
End Sub
