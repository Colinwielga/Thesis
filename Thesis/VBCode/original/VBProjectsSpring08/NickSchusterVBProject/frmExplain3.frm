VERSION 5.00
Begin VB.Form frmExplain3 
   BackColor       =   &H00800000&
   Caption         =   "Maximum Leg Press"
   ClientHeight    =   8460
   ClientLeft      =   3540
   ClientTop       =   1275
   ClientWidth     =   7605
   LinkTopic       =   "Form2"
   ScaleHeight     =   8460
   ScaleWidth      =   7605
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2760
      TabIndex        =   4
      Top             =   7200
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   3375
      Left            =   1560
      ScaleHeight     =   3315
      ScaleWidth      =   4155
      TabIndex        =   3
      Top             =   3600
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   "To Calculate:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   $"frmExplain3.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1695
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Your maximum leg press is the maximum amount of weight you can leg press one time."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   6855
   End
End
Attribute VB_Name = "frmExplain3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FITNESS ASSESSOR
'frmExplain3
'Nick Schuster
'March 26, 2008

'This form gives the user basic information about resting heart rate and how to calculate it.
Option Explicit
'To go back to the previous form
Private Sub cmdBack_Click()
frmExplain3.Hide
frmInfo.Show
End Sub
'To load the picture of proper leg press form
Private Sub Form_Load()
picResults.Picture = LoadPicture(App.Path & "\LegPress.gif")
End Sub


