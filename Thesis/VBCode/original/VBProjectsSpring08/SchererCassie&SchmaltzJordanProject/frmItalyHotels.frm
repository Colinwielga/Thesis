VERSION 5.00
Begin VB.Form frmItalyHotels 
   BackColor       =   &H000000C0&
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdhotel2 
      BackColor       =   &H00808000&
      Caption         =   "Hotel de Mediterranean"
      BeginProperty Font 
         Name            =   "Mathematica7"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   2535
   End
   Begin VB.CommandButton cmdhotel1 
      BackColor       =   &H00808000&
      Caption         =   "Castle de Isabella"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H0000C0C0&
      Caption         =   "Quit"
      Height          =   735
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H0000C0C0&
      Caption         =   "Back"
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label lbldirections 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Please choose which hotel you would like to stay at:"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   480
      Width           =   9255
   End
   Begin VB.Image Image2 
      Height          =   3375
      Left            =   6600
      Picture         =   "frmItalyHotels.frx":0000
      Top             =   1560
      Width           =   3945
   End
   Begin VB.Image Image1 
      Height          =   3210
      Left            =   720
      Picture         =   "frmItalyHotels.frx":32E6
      Top             =   1800
      Width           =   5250
   End
End
Attribute VB_Name = "frmItalyHotels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Travel Agency
'Form Name: frmItalyHotels
'Authors: Cassie Scherer and Jordan Schmaltz
'Date Written: 3/4/08
'Objective: The user can choose which hotel they would like to stay at while vacationing in Italy.


Private Sub cmdback_Click()

'Here we allow the user to return to the previous screen

frmItalyHotels.Hide
frmItalyStart.Show

End Sub

Private Sub cmdhotel1_Click()

'Here the user is directed to the screen where they make hotel reservations

frmItalyHotels.Hide
frmHotel.Show

End Sub

Private Sub cmdhotel2_Click()

'Here the user is directed to the screen where they make hotel reservations

frmItalyHotels.Hide
frmHotel.Show

End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub Form_Load()

'This code centers the form on computer screen upon loading

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

End Sub
