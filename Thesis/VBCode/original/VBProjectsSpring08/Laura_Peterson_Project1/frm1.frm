VERSION 5.00
Begin VB.Form frm1 
   Caption         =   "Main Screen"
   ClientHeight    =   10035
   ClientLeft      =   855
   ClientTop       =   1035
   ClientWidth     =   13425
   LinkTopic       =   "Form1"
   Picture         =   "frm1.frx":0000
   ScaleHeight     =   10035
   ScaleWidth      =   13425
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H00000080&
      Caption         =   "Enter Gallery"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro R"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7920
      Width           =   3255
   End
   Begin VB.Label lblName 
      BackColor       =   &H80000005&
      Caption         =   "Laura's Movie Gallery"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro R"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   10035
      Left            =   -960
      Picture         =   "frm1.frx":DAB5
      Top             =   0
      Width           =   15000
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnter_Click()
    frmAbout.Show
    
End Sub

