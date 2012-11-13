VERSION 5.00
Begin VB.Form frmFrozenFoods 
   BackColor       =   &H00000000&
   Caption         =   "Frozen Foods"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
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
      Top             =   5640
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
      Top             =   5640
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
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox txtSteak 
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
   Begin VB.TextBox txtShrimp 
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
   Begin VB.TextBox txtPizza 
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
   Begin VB.PictureBox picPizza 
      Height          =   1335
      Left            =   480
      Picture         =   "frmFrozenFoods.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   1995
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin VB.PictureBox picShrimp 
      Height          =   1215
      Left            =   480
      Picture         =   "frmFrozenFoods.frx":B5F2
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   3120
      Width           =   1575
   End
   Begin VB.PictureBox picSteak 
      Height          =   1455
      Left            =   480
      Picture         =   "frmFrozenFoods.frx":117B4
      ScaleHeight     =   1395
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   4560
      Width           =   1935
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
      Left            =   6840
      TabIndex        =   17
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label lblFrozenFoods 
      BackColor       =   &H00000000&
      Caption         =   "Frozen Foods"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   0
      Width           =   4935
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
   Begin VB.Label lblPizza 
      BackColor       =   &H00000000&
      Caption         =   "No frozen pizza provides the same quality at such an affordable price.  Red Baron pizza is perfect for any night! ($3.79/each)"
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
      Height          =   1095
      Left            =   3360
      TabIndex        =   6
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label lbShrimp 
      BackColor       =   &H00000000&
      Caption         =   $"frmFrozenFoods.frx":1AC76
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
      Height          =   1215
      Left            =   3360
      TabIndex        =   5
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Label lblSteak 
      BackColor       =   &H00000000&
      Caption         =   "This dinner combines both steak and mac n' cheese into one hearty portion that will fill anyone up! ($4.35/each)"
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
      Top             =   4560
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
Attribute VB_Name = "frmFrozenFoods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddToCart_Click()
pizza = txtPizza.Text
shrimp = txtShrimp.Text
steak = txtSteak.Text
frmFrozenFoods.Hide
frmHome.Show
'retrieving quantities from user
'then directing user back to home page
End Sub

Private Sub cmdCheckout_Click()
frmFrozenFoods.Hide
frmCheckOut.Show
'directing user to checkout form
End Sub

Private Sub cmdHome_Click()
frmFrozenFoods.Hide
frmHome.Show
'directing user to home form
End Sub

Private Sub CmdQuit_Click()
End
End Sub

