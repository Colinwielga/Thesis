VERSION 5.00
Begin VB.Form frmInfoPre 
   BackColor       =   &H00008000&
   Caption         =   "Steve Prefontaine"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Runners Page"
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label lblPre 
      BackColor       =   &H00008000&
      Caption         =   $"frmInfoPre.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   7095
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
   Begin VB.Image Image1 
      Height          =   6255
      Left            =   240
      Picture         =   "frmInfoPre.frx":01D9
      Top             =   0
      Width           =   3750
   End
End
Attribute VB_Name = "frmInfoPre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'goes back to runner's information page, text on form shows accomplishments of the runner'
Private Sub cmdBack_Click()
    frmRunners.Show
    frmInfoPre.Hide
End Sub

