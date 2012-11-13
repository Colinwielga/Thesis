VERSION 5.00
Begin VB.Form frmMsgBoxAviator 
   Caption         =   "Aviator"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Trends"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5520
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblAviator 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmMsgBoxAviator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdReturn_Click()
frmTrends.Show
frmMsgBoxAviator.Hide
End Sub
