VERSION 5.00
Begin VB.Form frmSuperDescription 
   BackColor       =   &H00800000&
   Caption         =   "Superpipe Description"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to main page"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00800000&
      Caption         =   $"frmSuperDescription.frx":0000
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
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmSuperDescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'description of superpipe competition with a link to the main page
Private Sub cmdReturn_Click()
    frmShaunWhite.Show
    frmSuperDescription.Hide
End Sub
