VERSION 5.00
Begin VB.Form RSX 
   BackColor       =   &H80000008&
   Caption         =   "Acura RSX"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   7800
      TabIndex        =   1
      Top             =   5280
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   240
      Picture         =   "frm8.frx":0000
      ScaleHeight     =   4635
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   240
      Width           =   8535
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Acura RSX"
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
      Top             =   5040
      Width           =   2295
   End
End
Attribute VB_Name = "RSX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    RSX.Hide
End Sub
