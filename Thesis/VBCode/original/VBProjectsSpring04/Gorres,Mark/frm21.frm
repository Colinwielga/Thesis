VERSION 5.00
Begin VB.Form Trailblazer 
   BackColor       =   &H80000008&
   Caption         =   "Chevy Trailblazer"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   3840
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   240
      Picture         =   "frm21.frx":0000
      ScaleHeight     =   3195
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Chevy Trailblazer"
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
      Top             =   3600
      Width           =   3495
   End
End
Attribute VB_Name = "Trailblazer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Trailblazer.Hide
End Sub

