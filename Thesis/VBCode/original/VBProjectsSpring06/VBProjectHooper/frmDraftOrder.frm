VERSION 5.00
Begin VB.Form frmDraftOrder 
   BackColor       =   &H00400000&
   Caption         =   "Form1"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10245
   FillColor       =   &H00800000&
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   7440
      Width           =   6255
   End
   Begin VB.PictureBox Picture1 
      Height          =   6975
      Left            =   240
      Picture         =   "frmDraftOrder.frx":0000
      ScaleHeight     =   6915
      ScaleWidth      =   9075
      TabIndex        =   0
      Top             =   120
      Width           =   9135
   End
End
Attribute VB_Name = "frmDraftOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'show draft order then navigate back to war room
Private Sub cmdBack_Click()
    frmDraftOrder.Hide
    frmWarRoom.Show
End Sub

