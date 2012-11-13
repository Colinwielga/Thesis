VERSION 5.00
Begin VB.Form frmwall 
   Caption         =   "Berlin Wall"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdmain 
      Caption         =   "Back to Berlin"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdquitber 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblwall 
      BackColor       =   &H8000000D&
      Caption         =   "Berlin Wall"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   $"frmwall.frx":0000
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   6615
   End
   Begin VB.Image Image3 
      Height          =   4050
      Left            =   0
      Picture         =   "frmwall.frx":012B
      Top             =   0
      Width           =   7275
   End
   Begin VB.Image Image1 
      Height          =   3735
      Left            =   120
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmwall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Tokyo, Berlin, Singapore- My Summer 2005 (Makihara_Kosuke.vbp)
'Form Name: Berlin Wall(frmwall.frm)
'Author: Kosuke Makihara
'Date Wrriten: 27 Oct 2005
'Ojectives:
'This form introduce some details of Berlin Wall


Private Sub cmdmain_Click()
frmwall.Hide
frmberlin.Show

End Sub

Private Sub cmdquitber_Click()
End

End Sub

