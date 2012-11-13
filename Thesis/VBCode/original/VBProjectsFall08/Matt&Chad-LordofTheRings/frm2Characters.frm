VERSION 5.00
Begin VB.Form frm2Characters 
   Caption         =   "Form1"
   ClientHeight    =   6720
   ClientLeft      =   855
   ClientTop       =   1515
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   Picture         =   "frm2Characters.frx":0000
   ScaleHeight     =   6720
   ScaleWidth      =   11160
   Begin VB.CommandButton cmdQuit 
      Caption         =   "End Your Journey"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Main Menu"
      Height          =   255
      Left            =   9960
      TabIndex        =   4
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Continue Your Adventure"
      Height          =   735
      Left            =   4560
      TabIndex        =   3
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton cmdAttributes 
      Caption         =   "Learn About Each Character's Attributes"
      Height          =   735
      Left            =   6480
      TabIndex        =   2
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdPictures 
      Caption         =   "Faces Of The Heroes"
      Height          =   735
      Left            =   4560
      TabIndex        =   1
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdBio 
      Caption         =   "Learn About Each Character"
      Height          =   735
      Left            =   2640
      TabIndex        =   0
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000010&
      BackStyle       =   0  'Transparent
      Caption         =   $"frm2Characters.frx":D764
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   960
      TabIndex        =   5
      Top             =   120
      Width           =   9975
   End
End
Attribute VB_Name = "frm2Characters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAttributes_Click()
    frm2Characters.Hide
    frm5Attributes.Show
End Sub

Private Sub cmdBio_Click()
    frm2Characters.Hide
    frm3Learn.Show
End Sub

Private Sub cmdChoose_Click()
    frm2Characters.Hide
    frm6Choose.Show
End Sub

Private Sub cmdPictures_Click()
    frm2Characters.Hide
    frm4Faces.Show
End Sub

Private Sub cmdQuit_Click()
    Quit
End Sub

Private Sub cmdReturn_Click()
    frm2Characters.Hide
    frm1HomePage.Show
End Sub
