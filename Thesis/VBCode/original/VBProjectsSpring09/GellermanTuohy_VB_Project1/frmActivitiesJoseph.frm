VERSION 5.00
Begin VB.Form frmActivitiesJoseph 
   BackColor       =   &H00FF00FF&
   Caption         =   "Activities"
   ClientHeight    =   9795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14160
   LinkTopic       =   "Form1"
   ScaleHeight     =   9795
   ScaleWidth      =   14160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "End Your Terrific Travel Experience With Johnnie Travel!!!! :'("
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   10440
      TabIndex        =   14
      Top             =   7200
      Width           =   2415
   End
   Begin VB.CommandButton cmdNext 
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
      Height          =   2295
      Left            =   240
      TabIndex        =   13
      Top             =   6120
      Width           =   9255
   End
   Begin VB.ComboBox cboCamping 
      Height          =   315
      ItemData        =   "frmActivitiesJoseph.frx":0000
      Left            =   3120
      List            =   "frmActivitiesJoseph.frx":0022
      TabIndex        =   12
      Top             =   4440
      Width           =   1575
   End
   Begin VB.ComboBox cboPizza 
      Height          =   315
      ItemData        =   "frmActivitiesJoseph.frx":0045
      Left            =   3240
      List            =   "frmActivitiesJoseph.frx":0067
      TabIndex        =   8
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdCamping 
      Caption         =   "Camping - Add"
      Height          =   615
      Left            =   5520
      TabIndex        =   5
      Top             =   4440
      Width           =   2535
   End
   Begin VB.ComboBox cboTicket 
      Height          =   315
      ItemData        =   "frmActivitiesJoseph.frx":008A
      Left            =   3240
      List            =   "frmActivitiesJoseph.frx":00AC
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdGaryspizza 
      Caption         =   "Gary's Pizza - Add"
      Height          =   615
      Left            =   5520
      TabIndex        =   1
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton cmdElpaso 
      Caption         =   "El Paso Bar - Add"
      Height          =   615
      Left            =   5520
      TabIndex        =   0
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Image Image3 
      Height          =   2100
      Left            =   9360
      Picture         =   "frmActivitiesJoseph.frx":00CF
      Top             =   4440
      Width           =   3165
   End
   Begin VB.Image Image2 
      Height          =   2700
      Left            =   10080
      Picture         =   "frmActivitiesJoseph.frx":AFAE
      Top             =   3120
      Width           =   3510
   End
   Begin VB.Image Image1 
      Height          =   2520
      Left            =   9360
      Picture         =   "frmActivitiesJoseph.frx":1591F
      Top             =   960
      Width           =   4170
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      Caption         =   "$15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   11
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "$10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   10
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000C0&
      Caption         =   "$20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   9
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblLunch 
      BackColor       =   &H00FF8080&
      Caption         =   "Food  (Pizza)      x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Camping"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lblElpaso 
      BackColor       =   &H00FF0000&
      Caption         =   "Concert Tickets    x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblActivities 
      BackColor       =   &H000080FF&
      Caption         =   "Activities in Saint Joseph, Minnesota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   240
      Width           =   9015
   End
End
Attribute VB_Name = "frmActivitiesJoseph"
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

Private Sub cmdCamping_Click()
    
    'Declare variable
    Dim Camping As Single
    
    'get the # of campers
    Camping = (CInt(cboCamping.Text) * 15)
    
    'running activity and checkout total
    ActivitiesTotal = ActivitiesTotal + Camping
    CheckoutTotal = CheckoutTotal + ActivitiesTotal
    
End Sub

Private Sub cmdElpaso_Click()

    'declare variable
    Dim Tickets As Single

    'get # of tickets
    Tickets = (CInt(cboTicket.Text) * 20) 'convert text to integer and calculate cost

    'running activity cost and checkout total
    ActivitiesTotal = ActivitiesTotal + Tickets
    CheckoutTotal = CheckoutTotal + ActivitiesTotal

End Sub

Private Sub cmdGaryspizza_Click()

    'declare variable
    Dim Pizzas As Single

    'get # of pizzas
    Pizzas = (CInt(cboPizza.Text) * 10) ' convert to integer then calculate cost

    'running activity cost and checkout total
    ActivitiesTotal = ActivitiesTotal + Pizzas
    CheckoutTotal = CheckoutTotal + ActivitiesTotal

End Sub

Private Sub cmdNext_Click()
    
    'hides current form and brings up checkout form
    frmActivitiesJoseph.Hide
    frmCheckout.Show
    
End Sub

Private Sub cmdQuit_Click()
    End         'ends the entire program when user presses it
End Sub

Private Sub Form_Load()
    'This code centers the form on computer screen upon loading

    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2

End Sub

