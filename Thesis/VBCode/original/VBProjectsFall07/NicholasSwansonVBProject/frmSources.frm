VERSION 5.00
Begin VB.Form frmSources 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sources"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4710
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ccmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label lblValidation 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmSources.frx":0000
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   1095
      Left            =   240
      TabIndex        =   6
      Top             =   4440
      Width           =   4335
   End
   Begin VB.Label lblVBCodeSource 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmSources.frx":0106
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "The SJU Athletics website provided two logos utilized by the program.  See <http://www.gojohnnies.com/>"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Label lblSJUMealPlans 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmSources.frx":01D6
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Label lblVBIndex 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmSources.frx":02B6
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Project Sources"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmSources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ccmdBack_Click()
'Returns user to Start Page
frmSources.Hide
End Sub
Private Sub Form_Load()
'centers form on computer screen upon loading
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
