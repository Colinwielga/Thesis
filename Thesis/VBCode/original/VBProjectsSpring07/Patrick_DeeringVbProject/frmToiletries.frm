VERSION 5.00
Begin VB.Form frmToiletries 
   BackColor       =   &H00000000&
   Caption         =   "Toiletries"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9285
   LinkTopic       =   "Form6"
   ScaleHeight     =   6930
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
      Left            =   3240
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
      Top             =   5880
      Width           =   1335
   End
   Begin VB.TextBox txtFloss 
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
   Begin VB.TextBox txtShampoo 
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
   Begin VB.TextBox txtToothpaste 
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
   Begin VB.PictureBox picToothpaste 
      Height          =   1095
      Left            =   600
      Picture         =   "frmToiletries.frx":0000
      ScaleHeight     =   1035
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.PictureBox picShampoo 
      Height          =   1215
      Left            =   720
      Picture         =   "frmToiletries.frx":55E6
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
   End
   Begin VB.PictureBox picFloss 
      Height          =   1695
      Left            =   600
      Picture         =   "frmToiletries.frx":D4E8
      ScaleHeight     =   1635
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   4560
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
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label lblToiletries 
      BackColor       =   &H00000000&
      Caption         =   "Toiletries"
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
      Left            =   3120
      TabIndex        =   9
      Top             =   0
      Width           =   3255
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
   Begin VB.Label lblToothpaste 
      BackColor       =   &H00000000&
      Caption         =   "Nobody makes better toothpaste than Crest! ($2.19/each)"
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
   Begin VB.Label lblShampoo 
      BackColor       =   &H00000000&
      Caption         =   "Suave's two-in-one shampoo/conditioner.  ($1.89/each)"
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
   Begin VB.Label lbFloss 
      BackColor       =   &H00000000&
      Caption         =   "Supersmiles floss - help keep your gums healthy! ($.49/each)"
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
Attribute VB_Name = "frmToiletries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddToCart_Click()
toothpaste = txtToothpaste.Text
shampoo = txtShampoo.Text
floss = txtFloss.Text
frmToiletries.Hide
frmHome.Show
'retrieving quantities from user
'then directing user back to home page
End Sub

Private Sub cmdCheckout_Click()
frmToiletries.Hide
frmCheckOut.Show
'direct user to CheckOut form
End Sub

Private Sub cmdHome_Click()
frmToiletries.Hide
frmHome.Show
'direct user to Home form
End Sub

Private Sub CmdQuit_Click()
End
End Sub

