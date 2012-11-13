VERSION 5.00
Begin VB.Form frmSlopeDescription 
   BackColor       =   &H00800000&
   Caption         =   "Slopestyle Description"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to main page"
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00800000&
      Caption         =   $"frmSlopeDescription.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmSlopeDescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this is a descriptive page on sopestyle competition with a link to the home page
Private Sub cmdReturn_Click()
    frmShaunWhite.Show
    frmSlopeDescription.Hide
End Sub
