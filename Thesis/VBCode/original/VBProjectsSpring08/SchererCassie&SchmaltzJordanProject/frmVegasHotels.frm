VERSION 5.00
Begin VB.Form frmVegasHotels 
   BackColor       =   &H0080FF80&
   Caption         =   "Form1"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12735
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   12735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H000080FF&
      Caption         =   "Quit"
      Height          =   975
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H000080FF&
      Caption         =   "Back"
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton cmdMGMgrand 
      BackColor       =   &H0000C000&
      Caption         =   "Stay at the MGM Grand!!"
      Height          =   975
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CommandButton cmdbellagio 
      BackColor       =   &H00FF8080&
      Caption         =   "Stay at the Bellagio!!"
      Height          =   975
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   3120
      Left            =   6000
      Picture         =   "frmVegasHotels.frx":0000
      Top             =   2040
      Width           =   5355
   End
   Begin VB.Image Image1 
      Height          =   3150
      Left            =   840
      Picture         =   "frmVegasHotels.frx":4E0E
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Label lbldirections 
      BackColor       =   &H0080FF80&
      Caption         =   "Please choose which hotel you want to stay at:"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   8775
   End
End
Attribute VB_Name = "frmVegasHotels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Travel Agency
'Form Name: frmVegasHotels
'Authors: Cassie Scherer and Jordan Schmaltz
'Date Written: 3/4/08
'Objective: The user can choose which hotel they would like to stay at while vacationing in Las Vegas.


Private Sub cmdback_Click()

'Here we are giving the user the option to return to the previous screen

frmVegasHotels.Hide
frmVegasStart.Show

End Sub

Private Sub cmdbellagio_Click()

'Here the user is directed to the form where they book their stay at the Bellagio

frmVegasHotels.Hide
frmHotel.Show

End Sub

Private Sub cmdMGMgrand_Click()

'Here the user is directed to the form where they book their stay at the MGM Grand

frmVegasHotels.Hide
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
