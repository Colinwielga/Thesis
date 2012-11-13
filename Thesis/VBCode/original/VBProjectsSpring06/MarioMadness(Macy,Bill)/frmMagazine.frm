VERSION 5.00
Begin VB.Form frmMagazine 
   Caption         =   "Nintendo Power Magazine Order Form"
   ClientHeight    =   9300
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   11868
   LinkTopic       =   "Form1"
   Picture         =   "frmMagazine.frx":0000
   ScaleHeight     =   9300
   ScaleWidth      =   11868
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox combosubscriptiontype 
      Height          =   288
      ItemData        =   "frmMagazine.frx":ED29
      Left            =   2760
      List            =   "frmMagazine.frx":ED36
      TabIndex        =   43
      Text            =   "Please Select the Type of Subscription"
      Top             =   7440
      Width           =   3972
   End
   Begin VB.ComboBox Combopaymenttype 
      Height          =   288
      ItemData        =   "frmMagazine.frx":ED89
      Left            =   2760
      List            =   "frmMagazine.frx":ED96
      TabIndex        =   42
      Text            =   "Please Select a Payment Type"
      Top             =   5400
      Width           =   2772
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit Your Subscription"
      Height          =   732
      Left            =   1920
      TabIndex        =   40
      Top             =   8160
      Width           =   2772
   End
   Begin VB.PictureBox PicMag3 
      Height          =   1212
      Left            =   1200
      Picture         =   "frmMagazine.frx":EDBB
      ScaleHeight     =   1164
      ScaleWidth      =   1764
      TabIndex        =   39
      Top             =   480
      Width           =   1812
   End
   Begin VB.ComboBox ComboSubscripYear 
      Height          =   288
      ItemData        =   "frmMagazine.frx":15F85
      Left            =   9600
      List            =   "frmMagazine.frx":15F92
      TabIndex        =   38
      Text            =   "Select a year"
      Top             =   1920
      Width           =   1452
   End
   Begin VB.ComboBox ComboExpYear 
      Height          =   288
      ItemData        =   "frmMagazine.frx":15FA8
      Left            =   9120
      List            =   "frmMagazine.frx":15FCD
      TabIndex        =   36
      Text            =   "Select a year"
      Top             =   6720
      Width           =   1452
   End
   Begin VB.ComboBox ComboSubscripMonth 
      Height          =   288
      ItemData        =   "frmMagazine.frx":16013
      Left            =   7800
      List            =   "frmMagazine.frx":1603B
      TabIndex        =   35
      Text            =   "Select a Month"
      Top             =   1920
      Width           =   1452
   End
   Begin VB.ComboBox ComboExpMonth 
      Height          =   288
      ItemData        =   "frmMagazine.frx":16080
      Left            =   7320
      List            =   "frmMagazine.frx":160A8
      TabIndex        =   34
      Text            =   "Select a Month"
      Top             =   6720
      Width           =   1452
   End
   Begin VB.ComboBox ComboCardType 
      Height          =   288
      ItemData        =   "frmMagazine.frx":160ED
      Left            =   7320
      List            =   "frmMagazine.frx":16100
      TabIndex        =   33
      Text            =   "Please Select a Credit Card"
      Top             =   5880
      Width           =   2772
   End
   Begin VB.TextBox txtCardNumber 
      Height          =   288
      Left            =   1920
      TabIndex        =   32
      Text            =   "0"
      Top             =   6600
      Width           =   3852
   End
   Begin VB.TextBox txtCardName 
      Height          =   288
      Left            =   1920
      TabIndex        =   31
      Text            =   "N/A"
      Top             =   6000
      Width           =   3852
   End
   Begin VB.TextBox txtCity 
      Height          =   288
      Left            =   1200
      TabIndex        =   30
      Top             =   3120
      Width           =   3492
   End
   Begin VB.TextBox txtPhone3 
      Height          =   288
      Left            =   9600
      TabIndex        =   29
      Top             =   4920
      Width           =   1212
   End
   Begin VB.TextBox txtPhone2 
      Height          =   288
      Left            =   8400
      TabIndex        =   27
      Top             =   4920
      Width           =   852
   End
   Begin VB.TextBox txtPhone1 
      Height          =   288
      Left            =   7200
      TabIndex        =   25
      Top             =   4920
      Width           =   852
   End
   Begin VB.TextBox txtEmail 
      Height          =   288
      Left            =   1680
      TabIndex        =   24
      Top             =   4560
      Width           =   3972
   End
   Begin VB.ComboBox ComboCountry 
      Height          =   288
      ItemData        =   "frmMagazine.frx":16139
      Left            =   4440
      List            =   "frmMagazine.frx":16146
      TabIndex        =   23
      Text            =   "Please Select a Country"
      Top             =   3840
      Width           =   2652
   End
   Begin VB.TextBox txtPostal 
      Height          =   288
      Left            =   1440
      TabIndex        =   22
      Top             =   3840
      Width           =   1692
   End
   Begin VB.TextBox txtPOBox 
      Height          =   288
      Left            =   9600
      TabIndex        =   21
      Top             =   2520
      Width           =   1092
   End
   Begin VB.ComboBox Combostate 
      Height          =   288
      ItemData        =   "frmMagazine.frx":16169
      Left            =   5640
      List            =   "frmMagazine.frx":16206
      TabIndex        =   20
      Text            =   "Please Select a State"
      Top             =   3120
      Width           =   2532
   End
   Begin VB.TextBox txtAddress 
      Height          =   288
      Left            =   1560
      TabIndex        =   18
      Top             =   2640
      Width           =   6492
   End
   Begin VB.TextBox txtName 
      Height          =   288
      Left            =   1560
      TabIndex        =   17
      Top             =   1920
      Width           =   4092
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to the main page"
      Height          =   732
      Left            =   6960
      TabIndex        =   0
      Top             =   8160
      Width           =   2652
   End
   Begin VB.Label lblMyname 
      BackStyle       =   0  'Transparent
      Caption         =   "By Bill Macy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   44
      Top             =   9000
      Width           =   1935
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   360
      X2              =   360
      Y1              =   360
      Y2              =   9600
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   11160
      X2              =   240
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   11280
      X2              =   11280
      Y1              =   0
      Y2              =   9600
   End
   Begin VB.Label lblSubscriptionType 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the Type of Subscription"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   600
      TabIndex        =   41
      Top             =   7200
      Width           =   1812
   End
   Begin VB.Label lblDash2 
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   8880
      TabIndex        =   37
      Top             =   6720
      Width           =   372
   End
   Begin VB.Label lblHyphen2 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   23.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   9360
      TabIndex        =   28
      Top             =   4800
      Width           =   252
   End
   Begin VB.Label lblHyphen1 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   23.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   8160
      TabIndex        =   26
      Top             =   4800
      Width           =   252
   End
   Begin VB.Label lblDash1 
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   9360
      TabIndex        =   19
      Top             =   1920
      Width           =   372
   End
   Begin VB.Label lblPaymentType 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select the Type of Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   600
      TabIndex        =   16
      Top             =   5040
      Width           =   2652
   End
   Begin VB.Label lblCardNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Card Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   600
      TabIndex        =   15
      Top             =   6480
      Width           =   1572
   End
   Begin VB.Label lblCardName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name on Card"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   600
      TabIndex        =   14
      Top             =   5760
      Width           =   1212
   End
   Begin VB.Label lblExpDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Expiration Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   6120
      TabIndex        =   13
      Top             =   6480
      Width           =   1332
   End
   Begin VB.Label lblCardType 
      BackStyle       =   0  'Transparent
      Caption         =   "Type of Card"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   6120
      TabIndex        =   12
      Top             =   5640
      Width           =   1092
   End
   Begin VB.Label lblsubscription 
      BackStyle       =   0  'Transparent
      Caption         =   "Date to Start the Subscription"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   6000
      TabIndex        =   11
      Top             =   1920
      Width           =   2172
   End
   Begin VB.Label lblPOBox 
      BackStyle       =   0  'Transparent
      Caption         =   "P.O. Box Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   8280
      TabIndex        =   10
      Top             =   2400
      Width           =   1332
   End
   Begin VB.Label lblPhoneNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   6120
      TabIndex        =   9
      Top             =   4920
      Width           =   1212
   End
   Begin VB.Label lblCity 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   600
      TabIndex        =   8
      Top             =   3120
      Width           =   372
   End
   Begin VB.Label lblPostal 
      BackStyle       =   0  'Transparent
      Caption         =   "Postal Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   600
      TabIndex        =   7
      Top             =   3600
      Width           =   972
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Caption         =   "Street Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   600
      TabIndex        =   6
      Top             =   2400
      Width           =   1452
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   600
      TabIndex        =   5
      Top             =   1920
      Width           =   732
   End
   Begin VB.Label lblState 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4920
      TabIndex        =   4
      Top             =   3120
      Width           =   492
   End
   Begin VB.Label lblCountry 
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3480
      TabIndex        =   3
      Top             =   3840
      Width           =   852
   End
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   600
      TabIndex        =   2
      Top             =   4320
      Width           =   1092
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nintendo Power Magazine"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   28.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   3120
      TabIndex        =   1
      Top             =   720
      Width           =   7452
   End
End
Attribute VB_Name = "frmMagazine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Mario Madness
'Form name: frmMagazine
'Author: Bill Macy
'Date Written: Tuesday March 14th, 2006
'Objective of form:  This form allows the user to enter all the information that would be needed for a magazine
                'subscription.  In all of the boxes the user will enter information to be submitted.  Nintendo
                'power is a magazine and this page allows the user to subscribe.  They can also return to the main
                'page
                
                
Option Explicit

Private Sub cmdreturn_Click()
    frmMagazine.Hide        'Hides the magazine page
    frmMain.Show        'shows the main page
End Sub

Private Sub cmdsubmit_Click()
    name1 = txtName.Text        'submits all the information entered in all the text boxes to variables defined in the program so the user can verify it in another form
    streetaddress = txtAddress.Text
    city = txtCity.Text
    postalcode = txtPostal.Text
    emailaddress = txtEmail.Text
    pobox = txtPOBox.Text
    phonenumber = txtPhone1.Text
    phonenumber2 = txtPhone2.Text
    phonenumber3 = txtPhone3.Text
    cardname = txtCardName.Text
    cardnumber = txtCardNumber.Text
    cardtype = ComboCardType
    expdate = ComboExpMonth
    expdate2 = ComboExpYear
    country = ComboCountry
    state = Combostate
    Date1 = ComboSubscripMonth
    date2 = ComboSubscripYear
    subscriptiontype = combosubscriptiontype
    paymenttype = Combopaymenttype
    frmcheck.Show       'shows the check form to verify everything
    frmMagazine.Hide        'hides the magazine page
End Sub
