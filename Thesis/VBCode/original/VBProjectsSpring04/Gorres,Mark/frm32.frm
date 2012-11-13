VERSION 5.00
Begin VB.Form SL500 
   BackColor       =   &H80000008&
   Caption         =   "M-B SL500"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   3600
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   240
      Picture         =   "frm32.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Mercedes-Benz SL500"
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
      Top             =   3360
      Width           =   4335
   End
End
Attribute VB_Name = "SL500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    SL500.Hide
End Sub
