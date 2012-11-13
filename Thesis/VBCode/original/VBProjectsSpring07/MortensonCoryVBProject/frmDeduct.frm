VERSION 5.00
Begin VB.Form frmDeduct 
   BackColor       =   &H00000000&
   Caption         =   "Schedule A"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHead 
      BackColor       =   &H0000FFFF&
      Caption         =   "Head of Household"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdSeparate 
      BackColor       =   &H0000FFFF&
      Caption         =   "Married Filing Seperate"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdMarried 
      BackColor       =   &H0000FFFF&
      Caption         =   "Married"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdClaimed 
      BackColor       =   &H0000FFFF&
      Caption         =   "Claimed Form "
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H0000FFFF&
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6120
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      Height          =   495
      Left            =   7560
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   17
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txtMisc 
      Height          =   495
      Left            =   7560
      TabIndex        =   9
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txtJob 
      Height          =   495
      Left            =   7560
      TabIndex        =   8
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtLosses 
      Height          =   495
      Left            =   7560
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtGifts 
      Height          =   495
      Left            =   7560
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtInterest 
      Height          =   495
      Left            =   7560
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtTaxes 
      Height          =   495
      Left            =   7560
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtMedical 
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdSingle 
      BackColor       =   &H0000FFFF&
      Caption         =   "Single Form"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Must Enter All Fields; (0 If Applicable.)"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3120
      TabIndex        =   24
      Top             =   720
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Return to:"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   23
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lbl8 
      BackColor       =   &H00000000&
      Caption         =   "Total Itemized Deductions"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2880
      TabIndex        =   16
      Top             =   6240
      Width           =   4335
   End
   Begin VB.Label lbl7 
      BackColor       =   &H00000000&
      Caption         =   "Other Miscellaneous Deductions"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2880
      TabIndex        =   15
      Top             =   5520
      Width           =   4335
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00000000&
      Caption         =   "Job Expenses and Certain Miscellaneous Deductions: (Exceeding 2% of Adjusted Gross Income)"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   2880
      TabIndex        =   14
      Top             =   4680
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Casualty and Theft/Losses"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2880
      TabIndex        =   13
      Top             =   4080
      Width           =   4335
   End
   Begin VB.Label lbl5 
      BackColor       =   &H00000000&
      Caption         =   "Gifts to Charity "
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2880
      TabIndex        =   12
      Top             =   3360
      Width           =   4335
   End
   Begin VB.Label lbl4 
      BackColor       =   &H00000000&
      Caption         =   "Interest You Paid: Home mortgage, Investment Interest, etc..."
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2880
      TabIndex        =   11
      Top             =   2640
      Width           =   4335
   End
   Begin VB.Label lbl3 
      BackColor       =   &H00000000&
      Caption         =   "Taxes You Paid: State and Local Income Taxes, Real estate, etc.."
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2880
      TabIndex        =   10
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00000000&
      Caption         =   "Medical and Dental Expenses: (Exceeding 7.5% of Adjusted Gross Income.)"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   2880
      TabIndex        =   3
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00000000&
      Caption         =   "Schedule A; Itemized Deductions"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "frmDeduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
    picResults.Cls
    
    'Reads the desired input into the variables
    Medical = txtMedical.Text
    Taxes = txtTaxes.Text
    Interest = txtInterest.Text
    Gifts = txtGifts.Text
    Losses = txtLosses.Text
    Job = txtJob.Text
    Misc = txtMisc.Text
    
    'Adds to find the total
    Itemized = Medical + Taxes + Interest + Gifts + Loss + Job + Misc
        picResults.Print FormatCurrency(Itemized)

    
End Sub


Private Sub cmdClaimed_Click()
    frmDeduct.Hide
    frmClaimed.Show
End Sub

Private Sub cmdHead_Click()
    frmDeduct.Hide
    frmHead.Show
End Sub

Private Sub cmdMarried_Click()
    frmDeduct.Hide
    frmMarried.Show
End Sub

Private Sub cmdSeparate_Click()
    frmDeduct.Hide
    frmSeperate.Show
End Sub

Private Sub cmdSingle_Click()
    frmDeduct.Hide
    frmSingle.Show
End Sub
