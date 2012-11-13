VERSION 5.00
Begin VB.Form frmCheckOut 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Checkout"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   Picture         =   "frmCheckOut.frx":0000
   ScaleHeight     =   6600
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEmailList 
      Caption         =   "Add your email to our monthly news letter with great deals and special offers!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      TabIndex        =   8
      Top             =   5280
      Width           =   2895
   End
   Begin VB.TextBox lblEmail 
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Text            =   "Email"
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox lblWork 
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Text            =   "Work Phone"
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox lblHome 
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1200
      TabIndex        =   5
      Text            =   "Home Phone"
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox lblName 
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Text            =   "Name on Credit Card"
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox lblExp 
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Text            =   "Exp"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox lblAddress 
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Text            =   "Billing Address"
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox lblCredit 
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Text            =   "Credit Card Number"
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton cmdVisa 
      Caption         =   "Show Reciept"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label lblCheckout 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks for shopping with us today.  Please fill out the following information .  "
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   360
      TabIndex        =   9
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "frmCheckOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dave and JB's ski and skate store
'frmFront
'Dave Lambert and John Brine
'Thursday March 23
'This form is used to get the billing and shipping information from the user to be used on other pages _
it also adds their email to a list

Private Sub cmdEmailList_Click()
    'displays a message box for adding email
    MsgBox "Thanks for addding your email keep an eye you for updates!", , "Email"
End Sub

Private Sub cmdVisa_Click()
    'allows the user to display all of the desired and necessary billing info
    'changes all of the input information into the variables
    Experation = lblExp
    Home = lblHome
    Work = lblWork
    Credit = lblCredit
    Address = lblAddress
    Email = lblEmail
    Name1 = lblName
    
    frmCheckOut.Visible = False ' changes to the checkout page
    frmReciept.Visible = True
End Sub

Private Sub Form_Load()
    'displays error message when you click on the pitcure box
    MsgBox "Please fill out all of the following information.", , "Do it!"
End Sub
