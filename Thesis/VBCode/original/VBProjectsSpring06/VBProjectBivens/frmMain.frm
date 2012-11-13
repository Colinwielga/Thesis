VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H000000FF&
   Caption         =   "Sexton Dining"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdLogout 
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3240
      TabIndex        =   12
      Top             =   6120
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel Order"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   11
      Top             =   6120
      Width           =   2775
   End
   Begin VB.CommandButton cmdPay 
      Caption         =   "Pay"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   10
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6240
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search and Sort"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   8
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdBakery 
      Caption         =   "Bakery"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   7
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton cmdDeliTacos 
      Caption         =   "Deli/Tacos"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3240
      TabIndex        =   6
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton cmdBeverages 
      Caption         =   "Beverages"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6240
      TabIndex        =   5
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton cmdPizza 
      Caption         =   "Pizza"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3240
      TabIndex        =   4
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CommandButton cmdSaladSoup 
      Caption         =   "Salad/Soup"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6240
      TabIndex        =   3
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CommandButton cmdGrill 
      Caption         =   "Grill"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   120
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   2655
      ScaleWidth      =   2655
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Paul Bivens"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   8400
      TabIndex        =   13
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Sexton Dining Price Computing Software"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sexton Dining Cash Register "\SextonDiningCashRegister.vpb"
'frmMain "\frmMain.frm"
'Paul Bivens
'March 22nd, 2006
'This form is used as a main form in which you are able to
'access the the rest of the forms.

Option Explicit
'Restarts your order by setting the sum equal to zero and gives you a message box
'confirming that the order has been cleared.
Private Sub cmdCancel_Click()
    Sum = 0
    MsgBox "Order Cleared", , "Order Cleared"
End Sub
'Takes you the grill form.
Private Sub cmdGrill_Click()
    frmGrill.Show
    frmMain.Hide
End Sub
'Brings you back to the entry form, restarts the timer and sets the counter used to
'determine the number of attempts at logging in to 0.
Private Sub cmdLogout_Click()
    frmEntry.Show
    frmMain.Hide
    LoginCounter = 0
    Timer1 = True
End Sub
'Takes you to the pay form.
Private Sub cmdPay_Click()
    frmPay.Show
    frmMain.Hide
End Sub
'Takes you to the pizza form.
Private Sub cmdPizza_Click()
    frmPizza.Show
    frmMain.Hide
End Sub
'Ends the program
Private Sub cmdQuit_Click()
    End
End Sub
'Takes you to the salad and soup form.
Private Sub cmdSaladSoup_Click()
    frmSaladSoup.Show
    frmMain.Hide
End Sub
'Takes you to the Bakery form.
Private Sub cmdBakery_Click()
    frmBakery.Show
    frmMain.Hide
    
End Sub
'Takes you to the deli and taco form.
Private Sub cmdDeliTacos_Click()
    frmDeliTacos.Show
    frmMain.Hide
    
End Sub
'Takes you to the beverage form.
Private Sub cmdBeverages_Click()
    frmBeverages.Show
    frmMain.Hide
    
End Sub
'Takes you to the search and sort form
Private Sub cmdSearch_Click()
    frmSearch.Show
    frmMain.Hide
End Sub
