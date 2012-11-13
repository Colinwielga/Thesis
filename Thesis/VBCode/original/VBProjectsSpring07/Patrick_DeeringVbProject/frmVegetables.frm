VERSION 5.00
Begin VB.Form frmVegetables 
   BackColor       =   &H00000000&
   Caption         =   "Vegetables"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9285
   LinkTopic       =   "Form4"
   ScaleHeight     =   6705
   ScaleWidth      =   9285
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
      Left            =   3360
      TabIndex        =   15
      Top             =   5760
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
      Left            =   6720
      TabIndex        =   14
      Top             =   5760
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
      Left            =   5040
      TabIndex        =   13
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox txtLettuce 
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
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox txtTomatoes 
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
   Begin VB.TextBox txtCarrots 
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
   Begin VB.PictureBox picCarrots 
      Height          =   1215
      Left            =   600
      Picture         =   "frmVegetables.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.PictureBox picTomatoes 
      Height          =   1215
      Left            =   600
      Picture         =   "frmVegetables.frx":C7C6
      ScaleHeight     =   1155
      ScaleWidth      =   1755
      TabIndex        =   1
      Top             =   3240
      Width           =   1815
   End
   Begin VB.PictureBox picLettuce 
      Height          =   1575
      Left            =   600
      Picture         =   "frmVegetables.frx":18F8C
      ScaleHeight     =   1515
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   4680
      Width           =   1815
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
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label lblVegetables 
      BackColor       =   &H00000000&
      Caption         =   "Vegetables"
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
      Left            =   3000
      TabIndex        =   9
      Top             =   0
      Width           =   3855
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
   Begin VB.Label lblCarrots 
      BackColor       =   &H00000000&
      Caption         =   "Our carrots are wonderful! They come pre-washed and ready to eat! ($.74/lb)"
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
   Begin VB.Label lblTomatoes 
      BackColor       =   &H00000000&
      Caption         =   "A wonderful addition to any salad, or eat them plain they're so sweet! ($.98/lb)"
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
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label lblLettuce 
      BackColor       =   &H00000000&
      Caption         =   "Our lettuce is as fresh as fresh can be.  All produce is brought in daily! ($.76/lb)"
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
      Top             =   4680
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
Attribute VB_Name = "frmVegetables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddToCart_Click()
carrots = txtCarrots.Text
tomatoes = txtTomatoes.Text
lettuce = txtLettuce.Text
frmVegetables.Hide
frmHome.Show
'retrieving quantities from user
'then directing user back to home page
End Sub

Private Sub cmdCheckout_Click()
frmVegetables.Hide
frmCheckOut.Show
'directs user to CheckOut form
End Sub

Private Sub cmdHome_Click()
frmVegetables.Hide
frmHome.Show
'directs user to Home form
End Sub

Private Sub CmdQuit_Click()
End
End Sub

