VERSION 5.00
Begin VB.Form frmLast 
   Caption         =   "Form1"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   720
      ScaleHeight     =   1395
      ScaleWidth      =   3915
      TabIndex        =   1
      Top             =   1680
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmLast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Picture1.Print Speed
End Sub
