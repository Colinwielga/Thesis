VERSION 5.00
Begin VB.Form frmPrograms 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Programs"
   ClientHeight    =   6795
   ClientLeft      =   2835
   ClientTop       =   2325
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   10260
   Begin VB.CommandButton cmdConvert 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Convert Your Money"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Find the Right Program for You!"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   5
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdSouthAfrica 
      BackColor       =   &H00FFFFC0&
      Caption         =   "South African Program"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdAsia 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Asian Programs"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdEurope 
      BackColor       =   &H00FFC0C0&
      Caption         =   "European Programs"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdSouthAmerica 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Chilean Program"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton cmdAustralia 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Australian Program"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   840
      Picture         =   "frmPrograms.frx":0000
      ScaleHeight     =   2355
      ScaleWidth      =   8475
      TabIndex        =   7
      Top             =   360
      Width           =   8535
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Study Abroad Programs"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   960
         TabIndex        =   8
         Top             =   1680
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmPrograms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'written 3/10/08 by Sammi and Erika

'the following 6 buttons bring the user to additional forms

Private Sub cmdAsia_Click()
    frmPrograms.Hide
    frmAsia.Show
End Sub

Private Sub cmdAustralia_Click()
    frmPrograms.Hide
    frmAustralia.Show
    
End Sub

Private Sub cmdConvert_Click()
    frmPrograms.Hide
    frmConvert.Show
End Sub

Private Sub cmdEurope_Click()
    frmPrograms.Hide
    frmEurope.Show
    
End Sub

Private Sub cmdFind_Click()
    frmPrograms.Hide
    frmFind.Show
End Sub

Private Sub cmdSouthAfrica_Click()
    frmPrograms.Hide
    frmSouthAfrica.Show
End Sub

Private Sub cmdSouthAmerica_Click()
    frmPrograms.Hide
    frmSouthAmerica.Show
    
End Sub

Private Sub cmdQuit_Click()
End

End Sub
