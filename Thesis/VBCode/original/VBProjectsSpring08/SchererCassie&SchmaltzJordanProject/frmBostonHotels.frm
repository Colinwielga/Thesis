VERSION 5.00
Begin VB.Form frmBostonHotels 
   BackColor       =   &H00808000&
   Caption         =   "Form1"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdhotel1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Boston Park Plaza Hotel"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton cmdhotel2 
      BackColor       =   &H000040C0&
      Caption         =   "Omni Parker House"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      Height          =   855
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00808080&
      Caption         =   "Back"
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   2895
      Left            =   5400
      Picture         =   "frmBostonHotels.frx":0000
      Top             =   1560
      Width           =   3750
   End
   Begin VB.Image Image1 
      Height          =   3150
      Left            =   360
      Picture         =   "frmBostonHotels.frx":7D8A
      Top             =   1440
      Width           =   4290
   End
   Begin VB.Label lbldirections 
      Caption         =   "Please choose which hotel you would like to stay at:"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   9135
   End
End
Attribute VB_Name = "frmBostonHotels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Travel Agency
'Form Name: frmBostonHotels
'Authors: Cassie Scherer and Jordan Schmaltz
'Date Written: 3/4/08
'Objective: The user can choose which hotel they would like to stay at while vacationing in Boston


Private Sub cmdback_Click()

'Here the user has the option to return to the previous screen

frmBostonHotels.Hide
frmBostonStart.Show

End Sub

Private Sub cmdhotel1_Click()

'Here the user goes to the form to book their hotel stay

frmBostonHotels.Hide
frmHotel.Show

End Sub

Private Sub cmdhotel2_Click()

'Here the user goes to the form to book their hotel stay

frmBostonHotels.Hide
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
