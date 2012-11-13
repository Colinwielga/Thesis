VERSION 5.00
Begin VB.Form frmExplanation 
   Caption         =   "Explanation of Method"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBP 
      Height          =   1455
      Left            =   1320
      Picture         =   "frmExplanation.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   5355
      TabIndex        =   5
      Top             =   3960
      Width           =   5415
   End
   Begin VB.CommandButton cmdReturnExp 
      Caption         =   "Return to Main Page"
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label lblBP 
      Caption         =   $"frmExplanation.frx":1061
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   3240
      Width           =   7215
   End
   Begin VB.Label lblPitchers 
      Caption         =   $"frmExplanation.frx":1170
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   7215
   End
   Begin VB.Label lblHitters 
      Caption         =   $"frmExplanation.frx":129D
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   7215
   End
   Begin VB.Label lblIntro 
      Caption         =   $"frmExplanation.frx":1512
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmExplanation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form contains an explanation of the methods used
Private Sub cmdReturnExp_Click()
    'Makes home page appear
    frmSearch.Visible = False
    frmExplanation.Visible = False
    frmHome.Visible = True
    frmRankings.Visible = False
End Sub
