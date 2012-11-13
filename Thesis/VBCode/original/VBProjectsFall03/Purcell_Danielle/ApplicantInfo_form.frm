VERSION 5.00
Begin VB.Form frmApplicantInfo 
   BackColor       =   &H00FF8080&
   Caption         =   "Form1"
   ClientHeight    =   9585
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   ScaleHeight     =   9585
   ScaleWidth      =   13350
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAppYrsAtPrevAddress 
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
      Left            =   10560
      TabIndex        =   35
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmdNextPage1 
      BackColor       =   &H00E0E0E0&
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
      Left            =   11400
      TabIndex        =   33
      Top             =   8280
      Width           =   1455
   End
   Begin VB.TextBox txtAppYrsPrevEmp 
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
      Left            =   11160
      TabIndex        =   32
      Top             =   7320
      Width           =   1935
   End
   Begin VB.TextBox txtYrsCurrentEmp 
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
      Left            =   10560
      TabIndex        =   31
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox txtAppCurrentEmp 
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
      TabIndex        =   29
      Top             =   5520
      Width           =   3615
   End
   Begin VB.TextBox txtAppPreviousEmployer 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2640
      TabIndex        =   26
      Top             =   7320
      Width           =   4815
   End
   Begin VB.TextBox txtAppNumberDependents 
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
      Left            =   10560
      TabIndex        =   24
      Top             =   6360
      Width           =   1815
   End
   Begin VB.TextBox txtAppGrossIncome 
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
      Left            =   3000
      TabIndex        =   22
      Top             =   6360
      Width           =   2535
   End
   Begin VB.TextBox txtAppPreviousZipCode 
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
      Left            =   10560
      TabIndex        =   20
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox txtAppPreviousState 
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
      TabIndex        =   19
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox txtAppPreviousCity 
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
      Top             =   4680
      Width           =   3135
   End
   Begin VB.TextBox txtAppPreviousAddress 
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
      TabIndex        =   14
      Top             =   3840
      Width           =   4455
   End
   Begin VB.TextBox txtAppYrsAtCurrentAddress 
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
      Left            =   10560
      TabIndex        =   12
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtAppName 
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
      TabIndex        =   10
      Top             =   1320
      Width           =   4215
   End
   Begin VB.TextBox txtAppCurrentZipCode 
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
      Left            =   10560
      TabIndex        =   9
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtAppCurrentState 
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
      Left            =   6240
      TabIndex        =   8
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox txtAppCurrentCity 
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
      TabIndex        =   7
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox txtAppCurrentAddress 
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
      Left            =   2040
      TabIndex        =   6
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Label lblAppYrsAtPrevAddress 
      BackColor       =   &H00E0E0E0&
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
      Left            =   7320
      TabIndex        =   34
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Label lblYrsCurrentEmp 
      BackColor       =   &H00E0E0E0&
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
      Left            =   7440
      TabIndex        =   30
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Label lblAppCurrentEmp 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   28
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label lblAppYrsPrevEmp 
      BackColor       =   &H00E0E0E0&
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
      Left            =   7800
      TabIndex        =   27
      Top             =   7440
      Width           =   3135
   End
   Begin VB.Label lblAppPreviousEmployer 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   25
      Top             =   7440
      Width           =   2295
   End
   Begin VB.Label lblAppNumberDependents 
      BackColor       =   &H00E0E0E0&
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
      Left            =   7800
      TabIndex        =   23
      Top             =   6480
      Width           =   2535
   End
   Begin VB.Label lblAppGrossIncome 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Gross Income Per Year (in dollars)"
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
      TabIndex        =   21
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Label lblAppPreviousZipCode 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   18
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblAppPreviousState 
      BackColor       =   &H00E0E0E0&
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
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label lblAppPreviousCity 
      BackColor       =   &H00E0E0E0&
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
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lblAppPreviousAddress 
      BackColor       =   &H00E0E0E0&
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
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label lblAppYrsAtAddress 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Years at Current Address"
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
      Left            =   7440
      TabIndex        =   11
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label lblAppCurrentZipCode 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9360
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblAppCurrentState 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5400
      TabIndex        =   4
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblAppCurrentCity 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   3
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblAppCurrentAddress 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblAppName 
      BackColor       =   &H00E0E0E0&
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
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblAppHeader 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Applicant General Information"
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
      Left            =   4560
      TabIndex        =   0
      Top             =   480
      Width           =   4815
   End
End
Attribute VB_Name = "frmApplicantInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdNextPage1_Click()
'Go on to the next page of the application.'

frmCoApplicantInfo.Show
frmApplicantInfo.Hide
End Sub

Private Sub Form_Load()
'When this form loads, it will be shown and then the message box will pop up.

frmApplicantInfo.Show
MsgBox ("You must fill in all fields.  If a certain question does not pertain to you, please enter 'X' for a field that would require text and '0' for a field that requires numbers")
End Sub
