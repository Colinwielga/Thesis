VERSION 5.00
Begin VB.Form frmXbox360Core 
   BackColor       =   &H004DD580&
   Caption         =   "Xbox 360 Core Unit"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   4440
      Left            =   3480
      Picture         =   "frmXbox360Core.frx":0000
      Top             =   480
      Width           =   3330
   End
End
Attribute VB_Name = "frmXbox360Core"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReturn_Click()
    frmXbox360Core.Hide
    frmConsoleInfo.Show
End Sub
