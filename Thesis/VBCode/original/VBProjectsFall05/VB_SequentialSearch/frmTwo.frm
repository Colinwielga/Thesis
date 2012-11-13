VERSION 5.00
Begin VB.Form frmTwo 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "show CTR"
      Height          =   735
      Left            =   2160
      TabIndex        =   1
      Top             =   1920
      Width           =   1815
   End
   Begin VB.PictureBox picResults2 
      Height          =   1095
      Left            =   480
      ScaleHeight     =   1035
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "frmTwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
picResults2.Print "the ctr is"; SequentialSearch.CTR
End Sub
