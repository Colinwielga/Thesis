VERSION 5.00
Begin VB.Form frmAtlanta 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form3"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   BeginProperty Font 
      Name            =   "@Batang"
      Size            =   36
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   8160
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Batang"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdATL 
      Caption         =   "Hotel Atlanta"
      BeginProperty Font 
         Name            =   "@Batang"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6120
      TabIndex        =   4
      Top             =   6360
      Width           =   3015
   End
   Begin VB.CommandButton cmdGF 
      Caption         =   "Girl Friend's House, Its FREE!"
      BeginProperty Font 
         Name            =   "@Batang"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   840
      TabIndex        =   3
      Top             =   6360
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   3090
      Left            =   5400
      Picture         =   "frmAtlanta.frx":0000
      Top             =   3240
      Width           =   4305
   End
   Begin VB.Label lblwhere 
      Alignment       =   2  'Center
      Caption         =   "Where To Stay?"
      Height          =   855
      Left            =   1200
      TabIndex        =   2
      Top             =   2160
      Width           =   8295
   End
   Begin VB.Image Image1 
      Height          =   2895
      Left            =   360
      Picture         =   "frmAtlanta.frx":2B782
      Top             =   3240
      Width           =   4305
   End
   Begin VB.Label lblPlayers 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Where the Players Play"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   1
      Top             =   1200
      Width           =   6015
   End
   Begin VB.Label lblAtlanta 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Welcome To Atlanta"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   7455
   End
End
Attribute VB_Name = "frmAtlanta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Atlanta page'
'This is another transition to another page promt'
'october 15th 2009'
'this is ateansition page from ATL to Hotel or No pay frm'
'Blake bauer'


'to hotel page'
Private Sub cmdATL_Click()
    frmAtlanta.Hide
    frmHotel.Show
End Sub
'quit Button'
Private Sub cmdEnd_Click()
    End
End Sub

'its free so it skips the hotel page'
Private Sub cmdGF_Click()
    frmAtlanta.Hide
    frmNoPay.Show
End Sub



Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
