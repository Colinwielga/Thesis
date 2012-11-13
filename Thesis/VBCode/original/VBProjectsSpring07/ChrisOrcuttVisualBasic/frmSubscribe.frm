VERSION 5.00
Begin VB.Form frmSubscribe 
   BackColor       =   &H000080FF&
   Caption         =   "Enter Subscirber Informatior"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   12450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSubscribe 
      Caption         =   "Subscribe"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   32
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   31
      Top             =   6720
      Width           =   2295
   End
   Begin VB.ComboBox ComboSelectSubscription 
      Height          =   315
      ItemData        =   "frmSubscribe.frx":0000
      Left            =   4680
      List            =   "frmSubscribe.frx":000D
      TabIndex        =   29
      Text            =   "Please select subcription type"
      Top             =   5880
      Width           =   4215
   End
   Begin VB.TextBox txtLastNum 
      Height          =   285
      Left            =   10920
      TabIndex        =   28
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtFirstNum 
      Height          =   285
      Left            =   10080
      TabIndex        =   26
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtAreaCode 
      Height          =   285
      Left            =   9240
      TabIndex        =   24
      Top             =   720
      Width           =   615
   End
   Begin VB.ComboBox ComboExpYear 
      Height          =   315
      ItemData        =   "frmSubscribe.frx":0071
      Left            =   9360
      List            =   "frmSubscribe.frx":0093
      TabIndex        =   21
      Text            =   "Select Year"
      Top             =   4920
      Width           =   1455
   End
   Begin VB.ComboBox ComboExpMonth 
      Height          =   315
      ItemData        =   "frmSubscribe.frx":00D3
      Left            =   7320
      List            =   "frmSubscribe.frx":00FB
      TabIndex        =   20
      Text            =   "Select Month"
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox txtCreditCardNumber 
      Height          =   285
      Left            =   2880
      TabIndex        =   18
      Top             =   4920
      Width           =   3615
   End
   Begin VB.TextBox txtCardHolderName 
      Height          =   315
      Left            =   2880
      TabIndex        =   16
      Top             =   4200
      Width           =   3615
   End
   Begin VB.ComboBox ComboCreditCard 
      Height          =   315
      ItemData        =   "frmSubscribe.frx":0140
      Left            =   7440
      List            =   "frmSubscribe.frx":0150
      TabIndex        =   14
      Text            =   "Select Credit Carrier"
      Top             =   4200
      Width           =   2295
   End
   Begin VB.ComboBox ComboPaymentMethod 
      Height          =   315
      ItemData        =   "frmSubscribe.frx":0183
      Left            =   4680
      List            =   "frmSubscribe.frx":0190
      TabIndex        =   12
      Text            =   "Please select payment type "
      Top             =   3480
      Width           =   2535
   End
   Begin VB.ComboBox ComboCountry 
      Height          =   315
      ItemData        =   "frmSubscribe.frx":01B5
      Left            =   6480
      List            =   "frmSubscribe.frx":01C2
      TabIndex        =   9
      Text            =   "Please choose a country"
      Top             =   2640
      Width           =   2415
   End
   Begin VB.ComboBox ComboState 
      Height          =   315
      ItemData        =   "frmSubscribe.frx":01E5
      Left            =   6480
      List            =   "frmSubscribe.frx":0282
      TabIndex        =   8
      Text            =   "Please select a state"
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox txtPostalCode 
      Height          =   285
      Left            =   4680
      TabIndex        =   7
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   3720
      TabIndex        =   6
      Top             =   1680
      Width           =   3855
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   4200
      TabIndex        =   5
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   4080
      TabIndex        =   4
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label lblSubscriptionMethod 
      BackColor       =   &H000080FF&
      Caption         =   "Select Subscription Method:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   30
      Top             =   5640
      Width           =   3015
   End
   Begin VB.Label lblDash2 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   27
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblDash1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   25
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblPhonenumber 
      BackColor       =   &H000080FF&
      Caption         =   "Phone #:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   23
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblSeparate 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   22
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label txtExpiration 
      BackColor       =   &H000080FF&
      Caption         =   "Expiration Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   19
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label lblCardNumber 
      BackColor       =   &H000080FF&
      Caption         =   "Credit Card #:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lblNameOnCard 
      BackColor       =   &H000080FF&
      Caption         =   "Card Holder's Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label lblSelectCardType 
      BackColor       =   &H000080FF&
      Caption         =   "Select Credit Card:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   13
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label lblPayment 
      BackColor       =   &H000080FF&
      Caption         =   "Choose Payment Method:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label lblDetails 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Please  Enter Subscriber Details and Credit Card Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   10
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label lblPostalCode 
      BackColor       =   &H000080FF&
      Caption         =   "Postal Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblCity 
      BackColor       =   &H000080FF&
      Caption         =   "City:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblAddress 
      BackColor       =   &H000080FF&
      Caption         =   "Address: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackColor       =   &H000080FF&
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   3480
      Left            =   120
      Picture         =   "frmSubscribe.frx":04C2
      Top             =   4320
      Width           =   2460
   End
   Begin VB.Image Image1 
      Height          =   3855
      Left            =   120
      Picture         =   "frmSubscribe.frx":3A35
      Top             =   120
      Width           =   2700
   End
   Begin VB.Image ImageSubscribe 
      Height          =   1950
      Left            =   9480
      Picture         =   "frmSubscribe.frx":74F2
      Top             =   1560
      Width           =   1950
   End
End
Attribute VB_Name = "frmSubscribe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form is open for user to use should (s)/he desire to subscribe
'to a membership service provided through the program.
Option Explicit
Private Sub cmdCancel_Click()
    frmSubscribe.Hide       'Hides Subscribe form
    frmSelectWant.Show        'Shows Register form
End Sub
'Data a user inputs is stored to these variables and
'is printed following the comlpetion of all forms.
'Reference was made to Bill Macy's Mario Madness VB project to
'to learn combo box technique
Private Sub cmdSubscribe_Click()
    Name1 = txtName.Text
    Address = txtAddress.Text
    City = txtCity.Text
    Postal = txtPostalCode.Text
    Country = ComboCountry
    State = ComboState
    Area = txtAreaCode.Text
    FirstNum = txtFirstNum.Text
    LastNum = txtLastNum.Text
    CardHolder = txtCardHolderName.Text
    CardNumber = txtCreditCardNumber.Text
    CardType = ComboCreditCard
    Exp1 = ComboExpMonth
    Exp2 = ComboExpYear
    SubscriptionType = ComboSelectSubscription
    PaymentMethod = ComboPaymentMethod
    frmSubscribe.Hide
    frmSubscriptionCheck.Show
End Sub

