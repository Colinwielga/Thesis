VERSION 5.00
Begin VB.Form frmXbox360Elite 
   BackColor       =   &H00000000&
   Caption         =   "Xbox 360 Elite Unit (Black)"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   6585
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
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   2820
      Left            =   3360
      Picture         =   "frmXbox360Elite.frx":0000
      Top             =   600
      Width           =   3750
   End
End
Attribute VB_Name = "frmXbox360Elite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReturn_Click()
    frmXbox360Elite.Hide
    frmConsoleInfo.Show
End Sub
