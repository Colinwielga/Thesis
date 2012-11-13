VERSION 5.00
Begin VB.Form Kia 
   BackColor       =   &H80000008&
   Caption         =   "Kia Rio"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form2"
   ScaleHeight     =   5565
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   240
      Picture         =   "frm1.frx":0000
      ScaleHeight     =   4155
      ScaleWidth      =   7995
      TabIndex        =   1
      Top             =   240
      Width           =   8055
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   7320
      TabIndex        =   0
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Kia Rio"
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
      Top             =   4560
      Width           =   1455
   End
End
Attribute VB_Name = "Kia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Kia.Hide
End Sub
