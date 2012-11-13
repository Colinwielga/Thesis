VERSION 5.00
Begin VB.Form Frontier 
   BackColor       =   &H80000008&
   Caption         =   "Nissan Frontier"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5355
   LinkTopic       =   "Form2"
   ScaleHeight     =   4710
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   3960
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   240
      Picture         =   "frm6.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Nissan Frontier"
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
      Top             =   3720
      Width           =   3015
   End
End
Attribute VB_Name = "Frontier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Frontier.Hide
End Sub
