VERSION 5.00
Begin VB.Form frmExplain4 
   BackColor       =   &H00800000&
   Caption         =   "Trunk Flexibility"
   ClientHeight    =   10005
   ClientLeft      =   3795
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   10005
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
      Top             =   8640
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   3975
      Left            =   1800
      ScaleHeight     =   3915
      ScaleWidth      =   3915
      TabIndex        =   3
      Top             =   4440
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Your trunk flexibility score is the number of inches you can extend your hands beyond your toes in a strait-leg sitting position."
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
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   6855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   $"frmExplain4.frx":0000
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
      Height          =   2895
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   6855
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
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "frmExplain4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FITNESS ASSESSOR
'frmExplain4
'Nick Schuster
'March 26, 2008

'This form gives the user basic information about resting heart rate and how to calculate it.
Option Explicit
'To go back to the previous form
Private Sub cmdBack_Click()
frmInfo.Show
frmExplain4.Hide
End Sub
'To load the picture of proper trunk flexion form
Private Sub Form_Load()
picResults.Picture = LoadPicture(App.Path & "\TrunkFlex.gif")
End Sub
