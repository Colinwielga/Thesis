VERSION 5.00
Begin VB.Form Frmsecondform 
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9390
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6765
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdnext 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Go to Main Page"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton cmdtotal 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton cmdcatcostume 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Buy Now"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmddogcostume 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Buy Now"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdfish 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Buy Now"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdcatbed 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Buy Now"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmddogbed 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Buy Now"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdfishbowl 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Buy Now"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   1335
   End
   Begin VB.PictureBox Picture4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   4
      Left            =   2640
      Picture         =   "Form1.frx":D91E2
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.PictureBox Picture4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   3
      Left            =   4440
      Picture         =   "Form1.frx":E0184
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   7
      Top             =   3600
      Width           =   1335
   End
   Begin VB.PictureBox Picture4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   2
      Left            =   2640
      Picture         =   "Form1.frx":E70BE
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   6
      Top             =   3600
      Width           =   1335
   End
   Begin VB.PictureBox Picture4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   1
      Left            =   840
      Picture         =   "Form1.frx":EDD48
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   5
      Top             =   3600
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4440
      Picture         =   "Form1.frx":F40F6
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.PictureBox Picture4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   0
      Left            =   840
      Picture         =   "Form1.frx":FB678
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FFC0FF&
      FillColor       =   &H00FF80FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   6360
      ScaleHeight     =   4395
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Cat Costume $30"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   20
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Dog Costume $30"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   19
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Turtle Lagoon $30"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   18
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cat bed $150"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Dog bed $150"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fish Bowl $30"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblmyshoppingcart 
      BackStyle       =   0  'Transparent
      Caption         =   "My Shopping Cart"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label lblstore 
      BackStyle       =   0  'Transparent
      Caption         =   "The Store"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Frmsecondform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'user can purchase things for his or her pet on this page


Option Explicit
Dim ctr As Single


Private Sub cmdcatbed_Click()
Dim catbed As Single

    catbed = 150
    ctr = ctr + catbed
    picresults.Print "Cat Bed:", FormatCurrency(catbed)
    
End Sub


Private Sub cmdcatcostume_Click()
Dim Catcostume As Single

    Catcostume = 30
    ctr = ctr + Catcostume
    picresults.Print "Cat Costume:", FormatCurrency(Catcostume)
    
End Sub

Private Sub cmddogbed_Click()
Dim dogbed As Single

    dogbed = 15
    ctr = ctr + dogbed
    picresults.Print "Dog Bed:", FormatCurrency(dogbed)

End Sub

Private Sub cmddogcostume_Click()
Dim dogcostume As Single

    dogcostume = 300
    ctr = ctr + dogcostume
    picresults.Print "Dog Costume:", FormatCurrency(dogcostume)

End Sub

Private Sub cmdfish_Click()
Dim turtlelagoon As Single

    turtlelagoon = 30
    ctr = ctr + turtlelagoon
    picresults.Print "Turtle lagoon:", FormatCurrency(turtlelagoon)
    
End Sub

Private Sub cmdfishbowl_Click()
Dim Fishbowl As Single
    
    Fishbowl = 30
    ctr = ctr + Fishbowl
    picresults.Print "Fish Bowl:", FormatCurrency(Fishbowl)

End Sub

Private Sub Cmdnext_Click()
Frmsecondform.Hide
frmfirstform.Show
End Sub

Private Sub cmdtotal_Click()
Dim total As Integer

total = ctr
picresults.Print "---------------------------------"
picresults.Print "Total:", FormatCurrency(total)

End Sub

Private Sub Form_Load()

End Sub
