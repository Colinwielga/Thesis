VERSION 5.00
Begin VB.Form frmPersonalInfo 
   BackColor       =   &H80000013&
   Caption         =   "Personal Info"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Frontpage"
      Height          =   615
      Left            =   5040
      TabIndex        =   15
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save and Continue"
      Height          =   615
      Left            =   1680
      TabIndex        =   14
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox txtStateZip 
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   3960
      Width           =   6255
   End
   Begin VB.TextBox txtApart 
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtSpouseLast 
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox txtLast 
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox txtAddress 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox txtSpouse 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label lblMyName 
      BackColor       =   &H80000013&
      Caption         =   "By Brent Mergen"
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
      Left            =   7440
      TabIndex        =   16
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label lblSateZip 
      Caption         =   "City, State, Zip code"
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   3720
      Width           =   3615
   End
   Begin VB.Label lblApt 
      Caption         =   "Apt. No."
      Height          =   255
      Left            =   5040
      TabIndex        =   11
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblSpouseLast 
      Caption         =   "Spouse Last"
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblLastName 
      Caption         =   "Last Name"
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblSpouseName 
      Caption         =   "If joint return, spouses first name and initial"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label lblAddress 
      Caption         =   "Address"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "frmPersonalInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'E-Z Tazes (Brent's E-ZTax Form VBProject.vbp)
'Personnal Info (frmPersonnalInfo)
'Brent Timothy Mergen
'24 March 2006
'Type in your information on this form for E-Z Taxes information, so they can process the tax return

Private Sub cmdReturn_Click()
    frmFrontpage.Show 'brings you to a new form
    frmPersonalInfo.Hide 'hides old form
End Sub

Private Sub cmdSave_Click()
    frmTaxInput.Show 'brings you to a new form
    frmPersonalInfo.Hide 'hides old form
End Sub

