VERSION 5.00
Begin VB.Form frmBrowse 
   BackColor       =   &H00004000&
   Caption         =   "The Campground!"
   ClientHeight    =   10815
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   ScaleHeight     =   10815
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFindHelper 
      Caption         =   "Find Wendy for more help."
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      TabIndex        =   18
      Top             =   5280
      Width           =   2655
   End
   Begin VB.CommandButton cmdLeave 
      Caption         =   "Leave the store with no purchases."
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8280
      TabIndex        =   17
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdGoToCheck 
      BackColor       =   &H000080FF&
      Caption         =   "Proceed to checkout."
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   5280
      MaskColor       =   &H000080FF&
      TabIndex        =   16
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton cmdFindWendy 
      Caption         =   "Find Wendy for more help."
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   15
      Top             =   11040
      Width           =   3495
   End
   Begin VB.PictureBox picMessKit 
      BackColor       =   &H80000009&
      Height          =   3975
      Left            =   5520
      Picture         =   "frmBrowse.frx":0000
      ScaleHeight     =   3915
      ScaleWidth      =   4275
      TabIndex        =   12
      Top             =   6240
      Width           =   4335
      Begin VB.Label lblMessKit 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "Mess Kit $14.99"
         BeginProperty Font 
            Name            =   "Orator Std"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2400
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdGoToCheckout 
      Caption         =   "Proceed to checkout."
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4080
      TabIndex        =   11
      Top             =   11040
      Width           =   3375
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Leave the store with no purchases."
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8160
      TabIndex        =   10
      Top             =   11160
      Width           =   1695
   End
   Begin VB.CommandButton cmdMessKit 
      Caption         =   "Click to add a mess kit to your cart."
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7680
      TabIndex        =   9
      Top             =   3360
      Width           =   2175
   End
   Begin VB.PictureBox picJacket 
      BackColor       =   &H80000009&
      Height          =   3495
      Left            =   240
      Picture         =   "frmBrowse.frx":5B1A
      ScaleHeight     =   3435
      ScaleWidth      =   3915
      TabIndex        =   7
      Top             =   6720
      Width           =   3975
      Begin VB.Label lblJacket 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Jacket   S-XL $124.99 XXL  $134.99"
         BeginProperty Font 
            Name            =   "Orator Std"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2400
         TabIndex        =   8
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox picTent 
      Height          =   3135
      Left            =   1560
      Picture         =   "frmBrowse.frx":8048
      ScaleHeight     =   3075
      ScaleWidth      =   3435
      TabIndex        =   5
      Top             =   3360
      Width           =   3495
      Begin VB.Label lblTent 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "8-person tent $199.99"
         BeginProperty Font 
            Name            =   "Orator Std"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1680
         TabIndex        =   6
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.PictureBox picSleepingBag 
      Height          =   3015
      Left            =   360
      Picture         =   "frmBrowse.frx":9640
      ScaleHeight     =   2955
      ScaleWidth      =   3315
      TabIndex        =   3
      Top             =   120
      Width           =   3375
      Begin VB.Label lblSleepingBag 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sleeping Bag $74.99"
         BeginProperty Font 
            Name            =   "Orator Std"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdJacket 
      Caption         =   "Click to add a jacket to your cart."
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7680
      TabIndex        =   2
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton cmdTent 
      Caption         =   "Click to add a tent to your cart."
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7680
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdSleepingBag 
      Caption         =   "Click to add a sleeping bag to your cart."
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7680
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblBrowse 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Check out our most popular products!"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2895
      Left            =   4440
      TabIndex        =   14
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFindHelper_Click()
'takes the user to the ProductSearch form using the Visible property
'also saves the subtotal for use at the Checkout form
frmBrowse.Hide
frmProductSearch.Show
Subtotal = SleepingBagSub + TentSub + JacketSub + MessKitSub
End Sub

Private Sub cmdGoToCheck_Click()
'takes the user to the Checkout form using the Visible property
'also saves the subtotal for use at the Checkout form
frmBrowse.Hide
frmCheckout.Show
Subtotal = SleepingBagSub + TentSub + JacketSub + MessKitSub
End Sub

Private Sub cmdJacket_Click()
'asks which size jacket the user wants using an inputbox, then adds to the counter
'for each jacket added to the cart
'adds a jacket to the subtotal and stores it for use in the Checkout form
Dim Size As String
RegJacket = 124.99
XXLJacket = 134.99
Size = InputBox("What size jacket do you want? (S, M, L, XL, XXL)?", "Choose a size.")
If LCase(Size) = "S" Or LCase(Size) = "M" Or LCase(Size) = "L" Or LCase(Size) = "XL" Or UCase(Size) = "S" Or UCase(Size) = "M" Or UCase(Size) = "L" Or UCase(Size) = "XL" Then
    RJCTR = RJCTR + 1
ElseIf LCase(Size) = "XXL" Or UCase(Size) = "XXL" Then
    XJCTR = XJCTR + 1
Else
    MsgBox "Invalid size."
End If
JacketSub = RJCTR * RegJacket + XJCTR * XXLJacket
End Sub

Private Sub cmdMessKit_Click()
'adds a mess kit to the subtotal and stores it for use in the Checkout form
MessKit = 14.99
MKCTR = MKCTR + 1
MessKitSub = MessKit * MKCTR
End Sub

Private Sub cmdLeave_Click()
'ends the program
End
End Sub

Private Sub cmdSleepingBag_Click()
'adds a sleeping bag to the subtotal and stores it for use in the Checkout form
SleepingBag = 74.99
SBCTR = SBCTR + 1
SleepingBagSub = SleepingBag * SBCTR
End Sub

Private Sub cmdTent_Click()
'adds a tent to the subtotal and stores it for use in the Checkout form
Tent = 199.99
TCTR = TCTR + 1
TentSub = Tent * TCTR
End Sub
