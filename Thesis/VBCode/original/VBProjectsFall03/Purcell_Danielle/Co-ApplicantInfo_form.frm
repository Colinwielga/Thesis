VERSION 5.00
Begin VB.Form frmCoApplicantInfo 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form2"
   ClientHeight    =   9585
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13425
   LinkTopic       =   "Form2"
   ScaleHeight     =   9585
   ScaleWidth      =   13425
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCoAppYrsPrevAddress 
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
      Left            =   10320
      TabIndex        =   36
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdPreviousPage1 
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
      Height          =   975
      Left            =   9840
      TabIndex        =   34
      Top             =   8280
      Width           =   1455
   End
   Begin VB.CommandButton cmdNextPage2 
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
      Height          =   975
      Left            =   11640
      TabIndex        =   33
      Top             =   8280
      Width           =   1455
   End
   Begin VB.TextBox txtCoAppYrsPrevEmp 
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
      Left            =   10320
      TabIndex        =   32
      Top             =   7320
      Width           =   1695
   End
   Begin VB.TextBox txtCoAppPrevEmp 
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
      Left            =   2400
      TabIndex        =   30
      Top             =   7320
      Width           =   3495
   End
   Begin VB.TextBox txtCoAppNumberDependents 
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
      Left            =   10320
      TabIndex        =   28
      Top             =   6480
      Width           =   1695
   End
   Begin VB.TextBox txtCoAppGrossIncome 
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
      Left            =   2880
      TabIndex        =   26
      Top             =   6480
      Width           =   2775
   End
   Begin VB.TextBox txtCoAppYrsCurrentEmp 
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
      Left            =   10320
      TabIndex        =   24
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtCoAppCurrentEmp 
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
      Left            =   2280
      TabIndex        =   22
      Top             =   5640
      Width           =   3975
   End
   Begin VB.TextBox txtCoAppPrevZipCode 
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
      Left            =   10320
      TabIndex        =   20
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox txtCoAppPrevState 
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
      Left            =   5760
      TabIndex        =   18
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox txtCoAppPrevCity 
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
      Left            =   960
      TabIndex        =   16
      Top             =   4800
      Width           =   3375
   End
   Begin VB.TextBox txtCoAppPrevAddress 
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
      Left            =   2280
      TabIndex        =   14
      Top             =   3960
      Width           =   4455
   End
   Begin VB.TextBox txtCoAppYrsCurrentAddress 
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
      Left            =   10320
      TabIndex        =   12
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtCoAppCurrentZipCode 
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
      Left            =   10320
      TabIndex        =   10
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox txtCoAppCurrentState 
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
      Left            =   5760
      TabIndex        =   8
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox txtCoAppCurrentCity 
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
      Left            =   1080
      TabIndex        =   6
      Top             =   3120
      Width           =   3375
   End
   Begin VB.TextBox txtCoAppCurrentAddress 
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
      TabIndex        =   4
      Top             =   2280
      Width           =   4695
   End
   Begin VB.TextBox txtCoAppName 
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
      Left            =   1200
      TabIndex        =   2
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label lblYrsPrevAddress 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Years At Previous Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   35
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Label lblCoAppYrsPrevEmp 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Years At Previous Employer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   31
      Top             =   7440
      Width           =   3015
   End
   Begin VB.Label lblCoAppPrevEmp 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Previous Employer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   7440
      Width           =   2055
   End
   Begin VB.Label lblCoAppNumberDependents 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Number of Dependents"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   27
      Top             =   6600
      Width           =   2775
   End
   Begin VB.Label lblCoAppGrossIncome 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gross Income per Year (in dollars)"
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
      Left            =   120
      TabIndex        =   25
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Label lblCoAppYrsCurrentEmp 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Years At Current Employer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   23
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Label lblCoAppCurrentEmp 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Current Employer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label lblCoAppPrevZipCode 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Zip Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   19
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lblCoAppPrevState 
      BackColor       =   &H00C0FFFF&
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   17
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label lblCoAppPrevCity 
      BackColor       =   &H00C0FFFF&
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label lblCoAppPrevAddress 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Previous Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label lblCoAppYrsCurrentAddress 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Years At Current Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   11
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label lblCoAppCurrentZipCode 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Zip Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   9
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label lblCoAppCurrentState 
      BackColor       =   &H00C0FFFF&
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblCoAppCurrentCity 
      BackColor       =   &H00C0FFFF&
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lblCoAppCurrentAddress 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Current Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblCoAppName 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblCoAppHeader 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Co-Applicant General Information"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "frmCoApplicantInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNextPage2_Click()
'Go on to the next page of the application.'
frmAssets.Show
frmCoApplicantInfo.Hide
End Sub

Private Sub cmdPreviousPage1_Click()
'Return to the previous page of the application.'
frmApplicantInfo.Show
frmCoApplicantInfo.Hide
End Sub

Private Sub Form_Load()
'When this form loads, the co-applicant form will be shown
'the applicant form will be hidden, and the message box will pop up.

frmCoApplicantInfo.Show
frmApplicantInfo.Hide
MsgBox ("You must fill in all fields.  If a certain question does not pertain to you, please enter 'X' for a field that would require text and '0' for a field that requires numbers")
End Sub

