VERSION 5.00
Begin VB.Form frmBadlandsActivity 
   BackColor       =   &H00004080&
   Caption         =   "Form1"
   ClientHeight    =   10275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14325
   LinkTopic       =   "Form1"
   ScaleHeight     =   10275
   ScaleWidth      =   14325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "End Your Awesome Travel Experience With Johnnie Travel!!! :'("
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   4680
      TabIndex        =   10
      Top             =   6360
      Width           =   2775
   End
   Begin VB.CommandButton cmdNextActivity 
      Caption         =   "Time For You To Check Out And Get On Your Way For Your Dream Vacation With Johnnie Travel!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   8160
      TabIndex        =   9
      Top             =   5280
      Width           =   5655
   End
   Begin VB.CommandButton cmdGeo 
      Caption         =   "Add"
      Height          =   495
      Left            =   10680
      TabIndex        =   6
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton cmdDriveloop 
      Caption         =   "Add"
      Height          =   495
      Left            =   10680
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.ComboBox cboPeople 
      Height          =   315
      ItemData        =   "frmBadlandsActivity.frx":0000
      Left            =   8280
      List            =   "frmBadlandsActivity.frx":0022
      TabIndex        =   4
      Text            =   "(people)"
      Top             =   3600
      Width           =   1935
   End
   Begin VB.ComboBox cboCars 
      Height          =   315
      ItemData        =   "frmBadlandsActivity.frx":0045
      Left            =   8280
      List            =   "frmBadlandsActivity.frx":0067
      TabIndex        =   2
      Text            =   "(cars)"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800080&
      Caption         =   "$22"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   8
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000080&
      Caption         =   "$15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   3060
      Left            =   360
      Picture         =   "frmBadlandsActivity.frx":008A
      Top             =   6360
      Width           =   3765
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "Geological Museum"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      TabIndex        =   3
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "34 Mile Driving Loop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Activities in the Badlands, South Dakota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   0
      Top             =   480
      Width           =   9255
   End
   Begin VB.Image Image2 
      Height          =   3930
      Left            =   240
      Picture         =   "frmBadlandsActivity.frx":74E3
      Top             =   1800
      Width           =   5250
   End
End
Attribute VB_Name = "frmBadlandsActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Vacation Planner
'Form Name: Destination
'Authors: Luke Gellerman and Tan Tuohy
'3/22/09
'This form allows the user to sign up for activites that are particular to this Location
'If the user's Location would have been different that theychose earlier in the program, then
'they would not have been brought to this page. The activities that the user signs up for are added to the
'CheckoutTotal that is used in the last form.

Option Explicit
Private Sub cmdDriveloop_Click()
    'declare variables
    Dim Driveloop As Single

    'get geo cost
    Driveloop = (CInt(cboCars.Text) * 22)

    'running activity cost and checkout total
    ActivitiesTotal = ActivitiesTotal + Driveloop
    CheckoutTotal = CheckoutTotal + ActivitiesTotal

End Sub

Private Sub cmdGeo_Click()

    'declare variables
    Dim Geo As Single

    'get geo cost
    Geo = (CInt(cboPeople.Text) * 15)

    'running activity cost and checkout total
    ActivitiesTotal = ActivitiesTotal + Geo
    CheckoutTotal = CheckoutTotal + ActivitiesTotal

End Sub

Private Sub cmdNextActivity_Click()
    'hides current form and brings up checkout form
    
    frmBadlandsActivity.Hide
    frmCheckout.Show
    
End Sub

Private Sub cmdQuit_Click()
    End         'ends program
End Sub

Private Sub Form_Load()
    'This code centers the form on computer screen upon loading

    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2

End Sub
