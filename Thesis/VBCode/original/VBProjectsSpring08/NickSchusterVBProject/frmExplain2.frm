VERSION 5.00
Begin VB.Form frmExplain2 
   BackColor       =   &H00800000&
   Caption         =   "Maximum Bench Press"
   ClientHeight    =   8265
   ClientLeft      =   3540
   ClientTop       =   1275
   ClientWidth     =   7590
   LinkTopic       =   "Form3"
   ScaleHeight     =   8265
   ScaleWidth      =   7590
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
      Top             =   7080
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   3255
      Left            =   1800
      ScaleHeight     =   3195
      ScaleWidth      =   3675
      TabIndex        =   3
      Top             =   3600
      Width           =   3735
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
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   $"frmExplain2.frx":0000
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
      Height          =   1815
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Your maximum bench press is the maximum amount of weight you can bench press one time."
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
Attribute VB_Name = "frmexplain2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FITNESS ASSESSOR
'frmExplain2
'Nick Schuster
'March 26, 2008

'This form gives the user basic information about resting heart rate and how to calculate it.
Option Explicit
'To go back to the previous form
Private Sub cmdBack_Click()
frmexplain2.Hide
frmInfo.Show
End Sub
'To load the picture of proper bench press form
Private Sub Form_Load()
picResults.Picture = LoadPicture(App.Path & "\BenchPress.gif")
End Sub
