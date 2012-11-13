VERSION 5.00
Begin VB.Form Expedition 
   BackColor       =   &H80000008&
   Caption         =   "Ford Expedition"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   8400
      TabIndex        =   1
      Top             =   6840
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   6375
      Left            =   120
      Picture         =   "frm23.frx":0000
      ScaleHeight     =   6315
      ScaleWidth      =   9195
      TabIndex        =   0
      Top             =   120
      Width           =   9255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Ford Expedition"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   6600
      Width           =   3135
   End
End
Attribute VB_Name = "Expedition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Expedition.Hide
End Sub

