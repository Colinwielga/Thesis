VERSION 5.00
Begin VB.Form Aveo 
   BackColor       =   &H80000008&
   Caption         =   "Chevy Aveo"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   3840
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   240
      Picture         =   "frm2.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   240
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Chevy Aveo"
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
      Width           =   2295
   End
End
Attribute VB_Name = "Aveo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Aveo.Hide
End Sub
