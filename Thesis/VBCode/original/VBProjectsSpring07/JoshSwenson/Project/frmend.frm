VERSION 5.00
Begin VB.Form frmend 
   BackColor       =   &H0000FF00&
   Caption         =   "TENNIS!!!"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Go Back"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   7200
      TabIndex        =   5
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdball2 
      Caption         =   "Ball"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   5880
      Width           =   615
   End
   Begin VB.CommandButton cmdball1 
      Caption         =   "Ball"
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   5880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1935
      Left            =   3960
      ScaleHeight     =   1875
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   5280
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   4815
      Left            =   0
      Picture         =   "frmend.frx":0000
      ScaleHeight     =   4755
      ScaleWidth      =   8355
      TabIndex        =   0
      Top             =   0
      Width           =   8415
   End
   Begin VB.Label Label1 
      Caption         =   "Try this fun tennis game!!"
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   4920
      Width           =   1815
   End
End
Attribute VB_Name = "frmend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'navigate the pages
Private Sub cmdback_Click()
frmend.Hide
frmsort.Show

End Sub
'makes one button disapear and the other reapear
Private Sub cmdball1_Click()
cmdball1.Visible = False
cmdball2.Visible = True

End Sub

Private Sub cmdball2_Click()
cmdball2.Visible = False
cmdball1.Visible = True

End Sub

Private Sub cmdquit_Click()
End
End Sub
