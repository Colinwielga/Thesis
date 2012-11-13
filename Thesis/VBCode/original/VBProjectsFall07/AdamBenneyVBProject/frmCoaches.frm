VERSION 5.00
Begin VB.Form frmCoaches 
   BackColor       =   &H000000FF&
   Caption         =   "COACHES PAGE"
   ClientHeight    =   10740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   15240
   ScaleWidth      =   25080
   Begin VB.CommandButton cmdBack1 
      Caption         =   "Back to Main Screen"
      Height          =   855
      Left            =   12600
      TabIndex        =   6
      Top             =   9720
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuiz 
      Caption         =   "Take a Coaches Quiz!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7440
      TabIndex        =   3
      Top             =   9360
      Width           =   4455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "MARK HELLENACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "THOMAS HAUGEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   7440
      TabIndex        =   4
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF80&
      Caption         =   $"frmCoaches.frx":0000
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   10560
      TabIndex        =   2
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   $"frmCoaches.frx":0221
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9975
      Left            =   3600
      TabIndex        =   1
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Image Image2 
      Height          =   5130
      Left            =   7440
      Picture         =   "frmCoaches.frx":0700
      Top             =   3000
      Width           =   2670
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "The Coaches"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
   Begin VB.Image Image1 
      Height          =   5505
      Left            =   240
      Picture         =   "frmCoaches.frx":3B0D
      Top             =   3000
      Width           =   2955
   End
End
Attribute VB_Name = "frmCoaches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form: Coaches page
    'This form displays some information about the coaches and their pictures.
    'It also has a button to send the user to the Quiz page.




Private Sub cmdBack1_Click()            'This button allows the user to navigate to main screen
frmFirst.Show
frmCoaches.Hide
frmPlayers.Hide
frmQuiz.Hide
End Sub

Private Sub cmdQuiz_Click()             'This button allows the user to navigate to the Quiz page.
frmQuiz.Show
frmCoaches.Hide
End Sub
