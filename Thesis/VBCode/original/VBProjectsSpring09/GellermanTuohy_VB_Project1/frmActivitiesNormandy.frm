VERSION 5.00
Begin VB.Form frmActivitiesNormandy 
   BackColor       =   &H0080FF80&
   Caption         =   "Normandy, France"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "End Your Magically Delicious Travel Experience With Johnnie Travel!! :'("
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   9960
      TabIndex        =   10
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Time For You To Check Out And Get On Your Way For Your Dream Vacation With Johnnie Travel!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   5640
      TabIndex        =   9
      Top             =   3240
      Width           =   4095
   End
   Begin VB.CommandButton cmdParagliding 
      Caption         =   "Add"
      Height          =   495
      Left            =   8880
      TabIndex        =   7
      Top             =   2400
      Width           =   1815
   End
   Begin VB.ComboBox cboParagliding 
      Height          =   315
      ItemData        =   "frmActivitiesNormandy.frx":0000
      Left            =   6960
      List            =   "frmActivitiesNormandy.frx":0022
      TabIndex        =   6
      Text            =   "# of people"
      Top             =   2400
      Width           =   1335
   End
   Begin VB.ComboBox cboMuseum 
      Height          =   315
      ItemData        =   "frmActivitiesNormandy.frx":0045
      Left            =   7080
      List            =   "frmActivitiesNormandy.frx":0067
      TabIndex        =   4
      Text            =   "# of people"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdMuseum 
      Caption         =   "Add"
      Height          =   495
      Left            =   8880
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   3285
      Left            =   240
      Picture         =   "frmActivitiesNormandy.frx":008A
      Top             =   3720
      Width           =   4890
   End
   Begin VB.Image Image1 
      Height          =   2340
      Left            =   240
      Picture         =   "frmActivitiesNormandy.frx":FBE0
      Top             =   1080
      Width           =   3120
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C000C0&
      Caption         =   "$75"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Para Gliding"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H000040C0&
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
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "D-Day Historical Museum"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Activities in Normandy, France"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   7575
   End
End
Attribute VB_Name = "frmActivitiesNormandy"
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
Private Sub cmdMuseum_Click()
    'declare variables
    Dim Museum As Single

    'get museum cost
    Museum = (CInt(cboMuseum.Text) * 15)

    'running activity cost and checkout cost
    ActivitiesTotal = ActivitiesTotal + Museum
    CheckoutTotal = CheckoutTotal + ActivitiesTotal

End Sub

Private Sub cmdParagliding_Click()

    'declare variables
    Dim Paragliding As Single

    'get museum cost
    Paragliding = (CInt(cboParagliding.Text) * 75)

    'running activity cost and checkout cost
    ActivitiesTotal = ActivitiesTotal + Paragliding
    CheckoutTotal = CheckoutTotal + ActivitiesTotal

End Sub

Private Sub cmdQuit_Click()
    End     'ends program
End Sub

Private Sub Command1_Click()
    'hides current form and brings up the checkout form

    frmActivitiesNormandy.Hide
    frmCheckout.Show
    
End Sub

Private Sub Form_Load()
    'This code centers the form on computer screen upon loading

    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2

End Sub
