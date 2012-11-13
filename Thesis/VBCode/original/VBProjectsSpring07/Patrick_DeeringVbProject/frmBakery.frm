VERSION 5.00
Begin VB.Form frmBakery 
   BackColor       =   &H00000000&
   Caption         =   "Bakery"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9285
   LinkTopic       =   "Form3"
   ScaleHeight     =   6630
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
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
      Left            =   3240
      TabIndex        =   16
      Top             =   5880
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
      Left            =   6600
      TabIndex        =   15
      Top             =   5880
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
      Left            =   4920
      TabIndex        =   14
      Top             =   5880
      Width           =   1335
   End
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
      TabIndex        =   13
      Top             =   0
      Width           =   255
   End
   Begin VB.TextBox txtPie 
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
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox txtLoaf 
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
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtBuns 
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
      Top             =   2040
      Width           =   975
   End
   Begin VB.PictureBox picPie 
      Height          =   1695
      Left            =   360
      Picture         =   "frmBakery.frx":0000
      ScaleHeight     =   1635
      ScaleWidth      =   2235
      TabIndex        =   3
      Top             =   4560
      Width           =   2295
   End
   Begin VB.PictureBox picLoaf 
      Height          =   1215
      Left            =   600
      Picture         =   "frmBakery.frx":C0B6
      ScaleHeight     =   1155
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.PictureBox picBuns 
      Height          =   1215
      Left            =   480
      Picture         =   "frmBakery.frx":136D8
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
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
      Left            =   6600
      TabIndex        =   17
      Top             =   5400
      Width           =   1815
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
      TabIndex        =   9
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblPie 
      BackColor       =   &H00000000&
      Caption         =   "Just like our bread, our apple pies are cooked daily! ($12/each)"
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
      TabIndex        =   8
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Label lblLoaf 
      BackColor       =   &H00000000&
      Caption         =   "Our famous Johnnie bread!  Need we say more? ($2.50/each)"
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
      TabIndex        =   7
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label lblBuns 
      BackColor       =   &H00000000&
      Caption         =   "Our buns are baked fresh every morning.  Buns are slice and sold in bags of 12.  ($3.25/bag)"
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
      Top             =   1920
      Width           =   3015
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
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
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
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   9240
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblBakery 
      BackColor       =   &H00000000&
      Caption         =   "Bakery"
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
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmBakery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddToCart_Click()
Dim cost As Single
buns = txtBuns.Text
bread = txtLoaf.Text
pie = txtPie.Text
'retrieving quantities for each product
frmBakery.Hide
frmHome.Show
'switching forms
End Sub

Private Sub cmdCheckout_Click()
frmBakery.Hide
frmCheckOut.Show
'moving to CheckOut form

End Sub

Private Sub cmdHome_Click()
frmBakery.Hide
frmHome.Show
'moving to Home form

End Sub

Private Sub CmdQuit_Click()
End
End Sub
'can quit program at anytime
