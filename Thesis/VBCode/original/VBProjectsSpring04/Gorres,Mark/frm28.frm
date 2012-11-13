VERSION 5.00
Begin VB.Form H2 
   BackColor       =   &H80000008&
   Caption         =   "Hummer H2"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   4080
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   240
      Picture         =   "frm28.frx":0000
      ScaleHeight     =   3435
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Hummer H2"
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
      Left            =   360
      TabIndex        =   2
      Top             =   3840
      Width           =   2535
   End
End
Attribute VB_Name = "H2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    H2.Hide
End Sub

