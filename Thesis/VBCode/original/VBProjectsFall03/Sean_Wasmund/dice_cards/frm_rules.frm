VERSION 5.00
Begin VB.Form frm_rules 
   BackColor       =   &H00FF8080&
   Caption         =   "Rules by Sean Wasmund"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_back 
      Caption         =   "Back"
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label lbl_rules 
      BackColor       =   &H00FF8080&
      Caption         =   $"frm_rules.frx":0000
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "frm_rules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_back_Click()
frm_rules.Hide
frm_bacc.Enabled = True
End Sub
