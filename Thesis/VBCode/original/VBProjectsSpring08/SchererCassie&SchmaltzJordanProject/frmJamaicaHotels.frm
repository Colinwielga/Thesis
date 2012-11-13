VERSION 5.00
Begin VB.Form frmJamaicaHotels 
   BackColor       =   &H00C00000&
   Caption         =   "Form1"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdhotel2 
      BackColor       =   &H008080FF&
      Caption         =   "The Tropical Retreat"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdhotel1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Beachside Bungalow"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   735
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H0000FFFF&
      Caption         =   "Back"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   2445
      Left            =   5280
      Picture         =   "frmJamaicaHotels.frx":0000
      Top             =   2040
      Width           =   4320
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   720
      Picture         =   "frmJamaicaHotels.frx":413D
      Top             =   1560
      Width           =   4275
   End
   Begin VB.Label lbldirerctions 
      Caption         =   "Please choose which hotel you want to stay at:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   9735
   End
End
Attribute VB_Name = "frmJamaicaHotels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Travel Agency
'Form Name: frmJamaicaHotels
'Authors: Cassie Scherer and Jordan Schmaltz
'Date Written: 3/4/08
'Objective: The user can choose which hotel they would like to stay at while vacationing in Jamaica.


Private Sub cmdback_Click()

'Here we allow the user to return to the previous page

frmJamaicaHotels.Hide
frmJamaicaStart.Show

End Sub

Private Sub cmdhotel1_Click()

'Here the user is directed to the screen where they make hotel reservations

frmJamaicaHotels.Hide
frmHotel.Show

End Sub

Private Sub cmdhotel2_Click()

'Here the user is directed to the screen where they make hotel reservations

frmJamaicaHotels.Hide
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
