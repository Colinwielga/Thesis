VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00800000&
   Caption         =   "Equity Sorter 1.0"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6750
   LinkTopic       =   "Form2"
   ScaleHeight     =   5850
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   2295
      Left            =   3360
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   2955
      TabIndex        =   4
      Top             =   3240
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   360
      Picture         =   "Form2.frx":3FF6
      ScaleHeight     =   3555
      ScaleWidth      =   2595
      TabIndex        =   3
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "CLICK ON A PICTURE TO EXIT"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   3840
      TabIndex        =   5
      Top             =   2400
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "(c) Fall 2003"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   2
      Top             =   1800
      Width           =   5535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   "Thanks For Using Equity Sorter 1.0!"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   600
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Created by John P. Elliott"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   1200
      Width           =   5535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_Click()
    End
End Sub

Private Sub Picture2_Click()
    End
End Sub
