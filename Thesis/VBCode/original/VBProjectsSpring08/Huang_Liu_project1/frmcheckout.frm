VERSION 5.00
Begin VB.Form frmcheckout 
   Caption         =   "Check Out"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   Picture         =   "frmcheckout.frx":0000
   ScaleHeight     =   7845
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   1800
      TabIndex        =   25
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   1800
      TabIndex        =   22
      Top             =   5520
      Width           =   3615
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   1800
      TabIndex        =   20
      Top             =   4920
      Width           =   2895
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   1800
      TabIndex        =   18
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   5280
      TabIndex        =   15
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   3000
      TabIndex        =   14
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   2520
      Width           =   4335
   End
   Begin VB.CommandButton cmdconfirm 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Confirm Order"
      Height          =   975
      Left            =   6840
      Picture         =   "frmcheckout.frx":4B83
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1920
      Width           =   4335
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back to the main page"
      Height          =   975
      Left            =   8640
      Picture         =   "frmcheckout.frx":4FA9
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtname 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Security Number"
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Online Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   23
      Top             =   3720
      Width           =   4455
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Card Nmber"
      Height          =   375
      Left            =   360
      TabIndex        =   21
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User Name"
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Card Type"
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Zipcode"
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Address line 2"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please Enter Your Information Below"
      BeginProperty Font 
         Name            =   "Berlin Sans FB"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Address line 1"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "frmcheckout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()

frmmain.Show
frmcheckout.Hide

End Sub

Private Sub cmdconfirm_Click()
MsgBox "Thanks for shopping", , "Thanks"

frmmain.Show
frmcheckout.Hide

End Sub

