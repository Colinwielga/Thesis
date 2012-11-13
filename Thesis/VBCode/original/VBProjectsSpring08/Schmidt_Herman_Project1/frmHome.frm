VERSION 5.00
Begin VB.Form frmHome 
   Caption         =   "Home"
   ClientHeight    =   10110
   ClientLeft      =   900
   ClientTop       =   690
   ClientWidth     =   13755
   ForeColor       =   &H0000FFFF&
   Icon            =   "frmHome.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmHome.frx":030A
   MousePointer    =   4  'Icon
   Picture         =   "frmHome.frx":0614
   ScaleHeight     =   10110
   ScaleWidth      =   13755
   Begin VB.CommandButton cmdGoToSummary 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Your Trip Summary!"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8640
      Width           =   3615
   End
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   5520
      Picture         =   "frmHome.frx":1D065
      ScaleHeight     =   3675
      ScaleWidth      =   2475
      TabIndex        =   11
      Top             =   2640
      Width           =   2535
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H8000000E&
      Height          =   3735
      Left            =   3000
      Picture         =   "frmHome.frx":2A434
      ScaleHeight     =   3675
      ScaleWidth      =   2475
      TabIndex        =   10
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8760
      Width           =   2175
   End
   Begin VB.PictureBox Picture6 
      Height          =   3735
      Left            =   10560
      Picture         =   "frmHome.frx":32780
      ScaleHeight     =   3675
      ScaleWidth      =   2355
      TabIndex        =   7
      Top             =   2640
      Width           =   2415
   End
   Begin VB.PictureBox Picture5 
      Height          =   3735
      Left            =   8040
      Picture         =   "frmHome.frx":383F6
      ScaleHeight     =   3675
      ScaleWidth      =   2475
      TabIndex        =   6
      Top             =   2640
      Width           =   2535
   End
   Begin VB.PictureBox Picture2 
      Height          =   3735
      Left            =   480
      Picture         =   "frmHome.frx":427F3
      ScaleHeight     =   3675
      ScaleWidth      =   2475
      TabIndex        =   5
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton cmdFlorida 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Florida"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton cmdColorado 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Colorado"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton cmdNewYork 
      BackColor       =   &H00FFFFC0&
      Caption         =   "New York"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton cmdCalifornia 
      BackColor       =   &H00FFFFC0&
      Caption         =   "California"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3240
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton cmdHawaii 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Hawaii"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   600
      MouseIcon       =   "frmHome.frx":4927A
      MousePointer    =   4  'Icon
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Travel Agency"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   720
      MouseIcon       =   "frmHome.frx":49584
      TabIndex        =   8
      Top             =   240
      Width           =   11775
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Travel Agency
'Form Name: Home
'Author: Taylor Herman & Mindy Schmidt
'Date Written: 3/23/08
'Objective: To inform people of 5 travel destinations in the U.S. to help them make
'           decisions on where they want to travel.

'Makes it so the user has to declare all of the variables.
Option Explicit

Private Sub cmdCalifornia_Click()
'When button is clicked, the Home Page hides and the California form shows.
frmCalifornia.Show
frmHome.Hide
End Sub


Private Sub cmdFlorida_Click()
'When button is clicked, the Home Page hides and the Florida form shows.
frmFlorida.Show
frmHome.Hide
End Sub

Private Sub cmdGoToSummary_Click()
'When button is clicked, the Home Page hides and the Summary form shows.
frmSummary.Show
frmHome.Hide
End Sub

Private Sub cmdHawaii_Click()
'When button is clicked, the Home Page hides and the Hawaii form shows.
frmHawaii.Show
frmHome.Hide
End Sub

Private Sub cmdNewYork_Click()
'When button is clicked, the Home Page hides and the New York form shows.
frmNewYork.Show
frmHome.Hide
End Sub

Private Sub cmdQuit_Click()
'Quits the program.
End
End Sub

Private Sub cmdColorado_Click()
'When button is clicked, the Home Page hides and the Colorado form shows.
frmColorado.Show
frmHome.Hide
End Sub



