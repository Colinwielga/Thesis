VERSION 5.00
Begin VB.Form frmDeli 
   BackColor       =   &H00000000&
   Caption         =   "Deli"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9315
   LinkTopic       =   "Form2"
   ScaleHeight     =   6690
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   16
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton cmdAddToCart 
      Caption         =   "Add to Cart"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdCheckout 
      Caption         =   "Check Out"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   14
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox txtSalmon 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   12
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox txtCheese 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   11
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox txtDeliMeat 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      Top             =   1920
      Width           =   975
   End
   Begin VB.PictureBox picDeliMeat 
      Height          =   1455
      Left            =   600
      Picture         =   "frmDeli.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.PictureBox picCheese 
      Height          =   1695
      Left            =   600
      Picture         =   "frmDeli.frx":7EAE
      ScaleHeight     =   1635
      ScaleWidth      =   1755
      TabIndex        =   1
      Top             =   3000
      Width           =   1815
   End
   Begin VB.PictureBox picSalmon 
      Height          =   1575
      Left            =   720
      Picture         =   "frmDeli.frx":10F08
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   4800
      Width           =   1575
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
      Left            =   6720
      TabIndex        =   17
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label lblDeli 
      BackColor       =   &H00000000&
      Caption         =   "Deli"
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
      Height          =   855
      Left            =   4080
      TabIndex        =   9
      Top             =   0
      Width           =   2535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   9240
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblItem 
      BackColor       =   &H00000000&
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00000000&
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblDeliMeat 
      BackColor       =   &H00000000&
      Caption         =   "The perfect deli meat platter for graduations or get-togethers!  Serves 25 people.  ($56/platter)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   3360
      TabIndex        =   6
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label lbCheese 
      BackColor       =   &H00000000&
      Caption         =   "One of the best cheese platters in town, a perfect compliment to the Deli Meat platter!  ($43/platter)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   3360
      TabIndex        =   5
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Label lblSalmon 
      BackColor       =   &H00000000&
      Caption         =   "Fresh salmon, perfect for Lent!  ($13.99/lb)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   3360
      TabIndex        =   4
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label lblQuantity 
      BackColor       =   &H00000000&
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "frmDeli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddToCart_Click()
deli = txtDeliMeat.Text
cheese = txtCheese.Text
salmon = txtSalmon.Text
frmDeli.Hide
frmHome.Show
'retrieving quantities from user
'then directing user back to home page
End Sub

Private Sub cmdCheckout_Click()
frmDeli.Hide
frmCheckOut.Show
'directing user to CheckOut form
End Sub

Private Sub cmdHome_Click()
frmDeli.Hide
frmHome.Show
'directing user to Home form
End Sub

Private Sub CmdQuit_Click()
End
End Sub
