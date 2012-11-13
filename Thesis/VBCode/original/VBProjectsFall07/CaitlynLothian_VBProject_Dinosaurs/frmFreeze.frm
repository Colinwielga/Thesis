VERSION 5.00
Begin VB.Form frmFreeze 
   Caption         =   "CHOMP!!!"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   Picture         =   "frmFreeze.frx":0000
   ScaleHeight     =   5085
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack6 
      BackColor       =   &H00004040&
      Caption         =   "Back to the main page"
      Height          =   735
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008080&
      Caption         =   $"frmFreeze.frx":F839
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmFreeze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBack6_Click()
    'Nagivates the user back to the main page
    frmFreeze.Visible = False
    frmLoad.Visible = True
End Sub
