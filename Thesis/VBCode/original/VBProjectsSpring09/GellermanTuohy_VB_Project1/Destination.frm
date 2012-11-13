VERSION 5.00
Begin VB.Form frmDestination 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   13170
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   19995
   LinkTopic       =   "Form1"
   ScaleHeight     =   13170
   ScaleWidth      =   19995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "End Your Wonderful Johnnie Travel Experience :'("
      Height          =   1215
      Left            =   8640
      TabIndex        =   10
      Top             =   11160
      Width           =   2295
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Page ====>"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   8880
      TabIndex        =   9
      Top             =   9360
      Width           =   1935
   End
   Begin VB.CommandButton cmdSaskatchewan 
      Caption         =   "Enjoy the Tourist Attractions of Saskatchewan!!!"
      Height          =   1575
      Left            =   9720
      TabIndex        =   8
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton cmdBadlands 
      Caption         =   "An Affordable And Fun Vacation!! Choose The BadLands!!!"
      Height          =   1575
      Left            =   7440
      TabIndex        =   7
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton cmdNormandy 
      Caption         =   "Yes!!!! I want to visit Normandy!!"
      Height          =   1575
      Left            =   9720
      TabIndex        =   6
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton cmdStjoe 
      Caption         =   "Pick St. Joe As Your Travel Destination"
      Height          =   1575
      Left            =   7320
      TabIndex        =   5
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      Caption         =   "Visit Saskatchewan for the beautiful outdoors and wonderful wildlife! Explore Mountains as far and tall as the eye can see!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11640
      TabIndex        =   4
      Top             =   11760
      Width           =   4095
   End
   Begin VB.Label Label4 
      BackColor       =   &H008080FF&
      Caption         =   $"Destination.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11760
      TabIndex        =   3
      Top             =   6120
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Vacationing in South Dakota is not only fun but affordable! Come see the Badlands of South Dakota!!"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   12120
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   $"Destination.frx":00D6
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
      Left            =   1200
      TabIndex        =   1
      Top             =   6240
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C000C0&
      Caption         =   "Here Are Our Awesome Destinations Offered By Johnnie Travel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   15135
   End
   Begin VB.Image Image4 
      Height          =   5325
      Left            =   1080
      Picture         =   "Destination.frx":0195
      Top             =   7080
      Width           =   7110
   End
   Begin VB.Image Image3 
      Height          =   5790
      Left            =   11640
      Picture         =   "Destination.frx":9184
      Top             =   7320
      Width           =   9000
   End
   Begin VB.Image Image2 
      Height          =   4980
      Left            =   11760
      Picture         =   "Destination.frx":1EFA4
      Top             =   1440
      Width           =   6225
   End
   Begin VB.Image Image1 
      Height          =   6615
      Left            =   2040
      Picture         =   "Destination.frx":27CEC
      Top             =   1320
      Width           =   7500
   End
End
Attribute VB_Name = "frmDestination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Vacation Planner
'Form Name: Destination
'Authors: Luke Gellerman and Tan Tuohy
'3/21/09
'Allow the user to select their destination and then go to the next page in the planning process
'The Location variable is saved on this form and it will determine what flights they can book and what activity form
'they will be taken to later in the program.

Option Explicit

Private Sub cmdBadlands_Click()
    'Location is publicly defined as a variable, and whichever destination the user selects, that destination becomes "Location"
    
    Location = "Badlands"
    
    MsgBox "Thank you for selecting " & Location & " as your travel destination!!"
    
    cmdNext.Enabled = True      'Enables the cmdNext button to bring the user to the next page of the program
    
End Sub

Private Sub cmdNext_Click()
    
    'brings the user to the next page in the process of planning their vacation
    frmDestination.Hide
    frmHotel.Show
    
End Sub

Private Sub cmdNormandy_Click()
    'Location is publicly defined as a variable, and whichever destination the user selects, that destination becomes "Location"
    
    Location = "Normandy"
    
    MsgBox " Thank you for selecting " & Location & " as your travel destination!!"
    
    cmdNext.Enabled = True      'Enables the cmdNext button to bring the user to the next page of the program
    
End Sub

Private Sub cmdQuit_Click()
    End     'ends the entrie program
End Sub

Private Sub cmdSaskatchewan_Click()
    'Location is publicly defined as a variable, and whichever destination the user selects, that destination becomes "Location"
    
    Location = "Saskatchewan"
    
    MsgBox " Thank you for selecting " & Location & " as your travel destination!!"
    
    cmdNext.Enabled = True      'Enables the cmdNext button to bring the user to the next page of the program
    
End Sub

Private Sub cmdStjoe_Click()
    'Location is publicly defined as a variable, and whichever destination the user selects, that destination becomes "Location"
    
    Location = "St. Joe"
    
    MsgBox "Thank you for selecting " & Location & " as your travel destination!!"
    
    cmdNext.Enabled = True      'Enables the cmdNext button to bring the user to the next page of the program
    
End Sub

Private Sub Form_Load()
    'This code centers the form on computer screen upon loading

    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2
    
End Sub

