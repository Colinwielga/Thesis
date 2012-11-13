VERSION 5.00
Begin VB.Form frmAssets 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form3"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9240
   LinkTopic       =   "Form3"
   ScaleHeight     =   8145
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNextPage3 
      Caption         =   "Next Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7680
      TabIndex        =   12
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton cmdPreviousPage2 
      Caption         =   "Previous Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5880
      TabIndex        =   11
      Top             =   7080
      Width           =   1455
   End
   Begin VB.PictureBox picTotalAssets 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2760
      ScaleHeight     =   1035
      ScaleWidth      =   2715
      TabIndex        =   10
      Top             =   5760
      Width           =   2775
   End
   Begin VB.CommandButton cmdCalculateAssets 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Calculate Value of Total Assets"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5880
      Width           =   1695
   End
   Begin VB.TextBox txtTotalOtherAssetsValue 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   8
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox txtTotalRealEstateValue 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   6
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox txtTotalVehicleValue 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   4
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox txtCashAmt 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   1560
      Width           =   5055
   End
   Begin VB.Label lblTotalOtherAssetsValue 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Total Value of Other Assets (in dollars)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label lblTotalRealEstateValue 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Total Value of Real Estate (in dollars)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label lblTotalVehicleValue 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Total Value of Vehicles (in dollars)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label lblCashAmt 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Amount of Cash (in dollars) "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblAssets 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Assets"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmAssets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form will be used to calculate the total assets of the applicant and his/her co-signer (if applicable).'
'The total assets will be used to determine the risk factor of borrowing the applicant money by comparing it to total liabilities.'

Private Sub cmdCalculateAssets_Click()
picTotalAssets.Cls
'Dimension variables as singles because they are all dollar amounts.'
Dim Cash As Single, VehicleValue As Single, RealEstate As Single, OtherValue As Single
Cash = txtCashAmt.Text
VehicleValue = txtTotalVehicleValue.Text
RealEstate = txtTotalRealEstateValue.Text
OtherValue = txtTotalOtherAssetsValue.Text

'Use the following formula to calculate total assets.

TotalAssets = Cash + VehicleValue + RealEstate + OtherValue
picTotalAssets.Print FormatCurrency(TotalAssets, 2)
End Sub

Private Sub cmdNextPage3_Click()
'Move on to next page in the application.'
frmLiabilities.Show
frmAssets.Hide
End Sub

Private Sub cmdPreviousPage2_Click()
'Go back a page in the application.'
frmCoApplicantInfo.Show
frmAssets.Hide
End Sub

Private Sub Form_Load()
'When this form loads, the current form will be shown, the previous form will be hidden,
'and the message box will appear.
frmAssets.Show
frmCoApplicantInfo.Hide
MsgBox ("You must fill in all fields.  If a certain question does not pertain to you, please enter 'X' for a field that would require text and '0' for a field that requires numbers")
End Sub

