VERSION 5.00
Begin VB.Form Gallardo 
   BackColor       =   &H80000008&
   Caption         =   "Lambo Gallardo"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   3240
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   240
      Picture         =   "frm33.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Lamborghini Gallardo"
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
      Top             =   3000
      Width           =   4335
   End
End
Attribute VB_Name = "Gallardo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Gallardo.Hide
End Sub
