VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "-+"
   ClientHeight    =   9930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbGoBack 
      Caption         =   "Go Back to Welcome Form"
      Height          =   615
      Left            =   480
      TabIndex        =   54
      Top             =   9000
      Width           =   1575
   End
   Begin VB.OptionButton radNo 
      BackColor       =   &H00FFC0C0&
      Caption         =   "No"
      Height          =   375
      Left            =   1680
      TabIndex        =   52
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton cmdCloseForm 
      Caption         =   "Close"
      Height          =   495
      Left            =   9240
      TabIndex        =   42
      Top             =   8760
      Width           =   1815
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   615
      Left            =   480
      TabIndex        =   41
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "Confirm Order"
      Height          =   615
      Left            =   7080
      TabIndex        =   40
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Calculate Total"
      Height          =   615
      Left            =   7080
      TabIndex        =   39
      Top             =   5760
      Width           =   1575
   End
   Begin VB.PictureBox picTotal 
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   9240
      ScaleHeight     =   2715
      ScaleWidth      =   1755
      TabIndex        =   38
      Top             =   5640
      Width           =   1815
   End
   Begin VB.ComboBox cmbSRQty 
      Height          =   315
      ItemData        =   "TheLaundryCo Form.frx":0000
      Left            =   5640
      List            =   "TheLaundryCo Form.frx":0013
      TabIndex        =   36
      Top             =   6240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton radYes 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Yes"
      Height          =   375
      Left            =   840
      TabIndex        =   35
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox txtFirstName 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtLastName 
      Height          =   285
      Left            =   4680
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtTelNo 
      Height          =   285
      Left            =   7920
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
   Begin VB.ComboBox cmbDCPantQty 
      Height          =   315
      ItemData        =   "TheLaundryCo Form.frx":0026
      Left            =   1800
      List            =   "TheLaundryCo Form.frx":0048
      TabIndex        =   31
      Top             =   4320
      Width           =   855
   End
   Begin VB.ComboBox cmbDCDressQty 
      Height          =   315
      ItemData        =   "TheLaundryCo Form.frx":006B
      Left            =   1800
      List            =   "TheLaundryCo Form.frx":008D
      TabIndex        =   30
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox txtDeliveryDate 
      Height          =   285
      Left            =   2160
      TabIndex        =   22
      Top             =   7320
      Width           =   2295
   End
   Begin VB.ComboBox cmbPSDressQty 
      Height          =   315
      ItemData        =   "TheLaundryCo Form.frx":00B0
      Left            =   9120
      List            =   "TheLaundryCo Form.frx":00D2
      TabIndex        =   20
      Top             =   4800
      Width           =   855
   End
   Begin VB.ComboBox cmbPSPantQty 
      Height          =   315
      ItemData        =   "TheLaundryCo Form.frx":00F5
      Left            =   9120
      List            =   "TheLaundryCo Form.frx":0117
      TabIndex        =   19
      Top             =   4320
      Width           =   855
   End
   Begin VB.ComboBox cmbPSShirtQty 
      Height          =   315
      ItemData        =   "TheLaundryCo Form.frx":013A
      Left            =   9120
      List            =   "TheLaundryCo Form.frx":015C
      TabIndex        =   18
      Top             =   3840
      Width           =   855
   End
   Begin VB.ComboBox cmbLDressQty 
      Height          =   315
      ItemData        =   "TheLaundryCo Form.frx":017F
      Left            =   5280
      List            =   "TheLaundryCo Form.frx":01A1
      TabIndex        =   16
      Top             =   4800
      Width           =   855
   End
   Begin VB.ComboBox cmbLPantQty 
      Height          =   315
      ItemData        =   "TheLaundryCo Form.frx":01C4
      Left            =   5280
      List            =   "TheLaundryCo Form.frx":01E6
      TabIndex        =   15
      Top             =   4320
      Width           =   855
   End
   Begin VB.ComboBox cmbLShirtQty 
      Height          =   315
      ItemData        =   "TheLaundryCo Form.frx":0209
      Left            =   5280
      List            =   "TheLaundryCo Form.frx":022B
      TabIndex        =   14
      Top             =   3840
      Width           =   855
   End
   Begin VB.ComboBox cmbDCShirtQty 
      Height          =   315
      ItemData        =   "TheLaundryCo Form.frx":024E
      Left            =   1800
      List            =   "TheLaundryCo Form.frx":0270
      TabIndex        =   12
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txtOrderDate 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblStainRate 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "$2/Stain"
      Height          =   195
      Left            =   2520
      TabIndex        =   53
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label lblLabelPSDress 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Dress"
      Height          =   195
      Left            =   8400
      TabIndex        =   51
      Top             =   4800
      Width           =   405
   End
   Begin VB.Label lblLabelPSPant 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Pant"
      Height          =   195
      Left            =   8400
      TabIndex        =   50
      Top             =   4440
      Width           =   330
   End
   Begin VB.Label lblLabelPSShirt 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Shirt"
      Height          =   195
      Left            =   8400
      TabIndex        =   49
      Top             =   3960
      Width           =   315
   End
   Begin VB.Label lblLabelLDress 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Dress"
      Height          =   195
      Left            =   4320
      TabIndex        =   48
      Top             =   4800
      Width           =   405
   End
   Begin VB.Label lblLabelLPant 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Pant"
      Height          =   195
      Left            =   4320
      TabIndex        =   47
      Top             =   4440
      Width           =   330
   End
   Begin VB.Label lblLabelLShirt 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Shirt"
      Height          =   195
      Left            =   4320
      TabIndex        =   46
      Top             =   3960
      Width           =   315
   End
   Begin VB.Label lblLabelDCDress 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Dress "
      Height          =   195
      Left            =   960
      TabIndex        =   45
      Top             =   4800
      Width           =   450
   End
   Begin VB.Label lblLabelDCPant 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Pant"
      Height          =   195
      Left            =   960
      TabIndex        =   44
      Top             =   4440
      Width           =   330
   End
   Begin VB.Label lblLabelDCShirt 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Shirt"
      Height          =   195
      Left            =   960
      TabIndex        =   43
      Top             =   3960
      Width           =   315
   End
   Begin VB.Label lblStainCount 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Please enter # of stains:"
      Height          =   195
      Left            =   3840
      TabIndex        =   37
      Top             =   6240
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label lblStainRemoval 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Stain Removal"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   34
      Top             =   5640
      Width           =   1350
   End
   Begin VB.Label lblDCPant 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "$3/Pant"
      Height          =   195
      Left            =   2760
      TabIndex        =   33
      Top             =   4320
      Width           =   585
   End
   Begin VB.Label lblDCDress 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "$4/Dress"
      Height          =   195
      Left            =   2760
      TabIndex        =   32
      Top             =   4800
      Width           =   660
   End
   Begin VB.Label lblPSDress 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "$2/Dress"
      Height          =   195
      Left            =   10320
      TabIndex        =   29
      Top             =   4800
      Width           =   660
   End
   Begin VB.Label lblPSPant 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "$1/Pant"
      Height          =   195
      Left            =   10320
      TabIndex        =   28
      Top             =   4320
      Width           =   585
   End
   Begin VB.Label lblPSShirt 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "$1/Shirt"
      Height          =   195
      Left            =   10320
      TabIndex        =   27
      Top             =   3840
      Width           =   570
   End
   Begin VB.Label lblLDress 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "$4.50/Dress"
      Height          =   195
      Left            =   6480
      TabIndex        =   26
      Top             =   4800
      Width           =   885
   End
   Begin VB.Label lblLPant 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "$2.50/Pant"
      Height          =   195
      Left            =   6480
      TabIndex        =   25
      Top             =   4320
      Width           =   810
   End
   Begin VB.Label lblLShirt 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "$1.50/Shirt"
      Height          =   195
      Left            =   6480
      TabIndex        =   24
      Top             =   3840
      Width           =   795
   End
   Begin VB.Label lblDCShirt 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "$2/Shirt"
      Height          =   195
      Left            =   2760
      TabIndex        =   23
      Top             =   3840
      Width           =   570
   End
   Begin VB.Label lblDeliveryDate 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Delivery Date"
      Height          =   195
      Left            =   720
      TabIndex        =   21
      Top             =   7320
      Width           =   960
   End
   Begin VB.Label lblPress 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Press"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   17
      Top             =   3360
      Width           =   450
   End
   Begin VB.Label lblLaundry 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Laundry"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   3360
      Width           =   765
   End
   Begin VB.Label lblDryCleaning 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Dry Cleaning"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   3360
      Width           =   1230
   End
   Begin VB.Label lblLaundryServices 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Laundry Services"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   720
      TabIndex        =   10
      Top             =   2760
      Width           =   2220
   End
   Begin VB.Label lblOrderDate 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Order Date"
      Height          =   195
      Left            =   840
      TabIndex        =   8
      Top             =   720
      Width           =   780
   End
   Begin VB.Label lblTelNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Telephone No."
      Height          =   195
      Left            =   6600
      TabIndex        =   7
      Top             =   1920
      Width           =   1065
   End
   Begin VB.Label lblLastName 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Last Name"
      Height          =   195
      Left            =   3720
      TabIndex        =   6
      Top             =   1920
      Width           =   765
   End
   Begin VB.Label lblFirstName 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "First Name"
      Height          =   195
      Left            =   720
      TabIndex        =   5
      Top             =   1920
      Width           =   750
   End
   Begin VB.Label lblCustInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Customer Information"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   3045
   End
   Begin VB.Label lblNewOrder 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "New Order"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Width           =   2205
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbGoBack_Click()
Form1.Hide
frmWelcome.Show
End Sub

Private Sub cmdClear_Click()
'Clearing all fields
CustFirstName = ""
CustLastName = ""
CustTelNo = ""
txtFirstName.Text = ""
txtLastName.Text = ""
txtTelNo.Text = ""
cmbDCShirtQty.Text = ""
cmbDCPantQty.Text = ""
cmbDCDressQty.Text = ""
cmbLShirtQty.Text = ""
cmbLPantQty.Text = ""
cmbLDressQty.Text = ""
cmbPSShirtQty.Text = ""
cmbPSPantQty.Text = ""
cmbPSDressQty.Text = ""
cmbSRQty.Text = ""
radNo.Value = True
txtDeliveryDate.Text = ""
picTotal.Cls
End Sub

Private Sub cmdCloseForm_Click()
'Going to frmWelcome
Form1.Hide
frmWelcome.Show
Call cmdClear_Click
End Sub

Private Sub cmdConfirm_Click()
'Every order must have a customer telephone number and delivery date
If txtTelNo.Text <> "" And txtDeliveryDate.Text <> "" Then
'Writing customer information in a text file
    Open App.Path & "\TheLaundryCo.txt" For Append As #1 'Appending to existing customer file
    Print #1, txtFirstName.Text; Tab(15); ", "; txtLastName.Text; Tab(30); ", "; txtTelNo.Text; Tab(45); ","; txtOrderDate.Text; Tab(60); ", "; OrderTotal; Tab(75); ", "; txtDeliveryDate.Text
    Close #1
    MsgBox ("The order has been confirmed.")
    'clearing all the fields so the next order can be entered, information retrived from http://en.wikiversity.org/wiki/Functions_and_Subroutines_in_VB6
    Call cmdClear_Click
Else: MsgBox ("Please check if both customer telephone number and delivery date have been entered.")
End If


End Sub

Private Sub cmdTotal_Click()
picTotal.Cls
Dim DCSubtotal As Single 'This holds the subtotal for Drycleaning Services
Dim LSubtotal As Single 'This holds the subtotal for Laundry Services
Dim PsSubtotal As Single 'This holds the subtotal for Press Services
Dim Subtotal As Single 'This holds the subtotal of all services
Dim StainSubtotal As Single 'This holds the subtotal of all Stain
Dim Tax As Single 'Total Tax
Dim Total As Single 'Total of all services plus tax minus discount
Dim Discount As Single
Discount = 0 'initializing Discount as zero

'Drycleaning subtotal for each type of garment
Dim DCShirt As Single, DCPant As Single, DCDress As Single

If cmbDCShirtQty.Text <> "" Then
    DCShirt = cmbDCShirtQty * 2
End If

If cmbDCPantQty.Text <> "" Then
    DCPant = cmbDCPantQty * 3
End If

If cmbDCDressQty.Text <> "" Then
    DCDress = cmbDCDressQty * 4
End If
'Subtotal for Drycleaning
DCSubtotal = DCShirt + DCPant + DCDress


'Laundry subtotal for each type of garment
Dim LShirt As Single, LPant As Single, LDress As Single

If cmbLShirtQty.Text <> "" Then
    LShirt = cmbLShirtQty * 1.5
End If

If cmbLPantQty.Text <> "" Then
    LPant = cmbLPantQty * 2.5
End If

If cmbLDressQty.Text <> "" Then
    LDress = cmbLDressQty * 4.5
End If
'Subtotal for Laundry
LSubtotal = LShirt + LPant + LDress


'Press subtotal for each type of garment
Dim PSShirt As Single, PSPant As Single, PSDress As Single

If cmbPSShirtQty.Text <> "" Then
    PSShirt = cmbPSShirtQty * 1
End If

If cmbPSPantQty.Text <> "" Then
    PSPant = cmbPSPantQty * 1
End If

If cmbPSDressQty.Text <> "" Then
    PSDress = cmbPSDressQty * 2
End If
'Subtotal for Press
PsSubtotal = PSShirt + PSPant + PSDress


'Subtotal for Stain Removal
If cmbSRQty.Text <> "" Then
    StainSubtotal = cmbSRQty.Text * 2
End If
 'Calculates subtotal for the order
Subtotal = DCSubtotal + LSubtotal + PsSubtotal + StainSubtotal


'Tax calculation prior to discount
Tax = Subtotal * 0.07

'Calculating discount--If subtotal is greater than $50 customer is awarded 10% discount
If Subtotal > 50 Then
Discount = Subtotal * 0.1
End If

'Calculating Total
Total = Subtotal + Tax - Discount
OrderTotal = Total
'Printing in a picturebox
picTotal.Print "Subtotal:", FormatCurrency(Subtotal)
picTotal.Print "Tax:", FormatCurrency(Tax)
picTotal.Print "Discount:", FormatCurrency(Discount)
picTotal.Print "--------------------------------------"
picTotal.Print "Total", FormatCurrency(Total)




End Sub

Private Sub Form_Load()
txtOrderDate.Text = DateValue(Now)
'setting the found customer values from customer search on Welcome Form
txtFirstName.Text = CustFirstName
txtLastName.Text = CustLastName
txtTelNo.Text = CustTelNo
End Sub

Private Sub radNo_Click()
cmbSRQty.Clear
If radNo.Value = True Then
    lblStainCount.Visible = False
    cmbSRQty.Visible = False
End If

End Sub

Private Sub radYes_Click()
If radYes.Value = True Then
    lblStainCount.Visible = True
    cmbSRQty.Visible = True
End If

    
End Sub

