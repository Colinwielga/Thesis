VERSION 5.00
Begin VB.Form frmExplain1 
   BackColor       =   &H00800000&
   Caption         =   "Resting Heart Rate"
   ClientHeight    =   7635
   ClientLeft      =   3540
   ClientTop       =   1440
   ClientWidth     =   7695
   FillColor       =   &H80000017&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   7695
   Begin VB.PictureBox picResults 
      Height          =   2775
      Left            =   1920
      ScaleHeight     =   2715
      ScaleWidth      =   3795
      TabIndex        =   4
      Top             =   3240
      Width           =   3855
   End
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
      Left            =   2880
      TabIndex        =   0
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   $"frmExplain1.frx":0000
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
      Height          =   1335
      Left            =   360
      TabIndex        =   3
      Top             =   1800
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
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Your resting heart rate is the number of times your heart beats in one minute while you are at rest."
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
      TabIndex        =   1
      Top             =   360
      Width           =   6855
   End
End
Attribute VB_Name = "frmExplain1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FITNESS ASSESSOR
'frmExplain1
'Nick Schuster
'March 26, 2008

'This form gives the user basic information about resting heart rate and how to calculate it.
Option Explicit
'To go back to the previous form
Private Sub cmdBack_Click()
frmExplain1.Hide
frmInfo.Show
End Sub
'To load the picture of how to measure heart rate
Private Sub Form_Load()
picResults.Picture = LoadPicture(App.Path & "\HeartRate.gif")
End Sub
