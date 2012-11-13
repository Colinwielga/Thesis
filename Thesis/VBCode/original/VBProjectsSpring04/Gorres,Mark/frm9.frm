VERSION 5.00
Begin VB.Form Jetta 
   BackColor       =   &H80000008&
   Caption         =   "Volkswagen Jetta"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   4680
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   240
      Picture         =   "frm9.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Volkswagen Jetta"
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
      Top             =   4440
      Width           =   3375
   End
End
Attribute VB_Name = "Jetta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Jetta.Hide
End Sub
