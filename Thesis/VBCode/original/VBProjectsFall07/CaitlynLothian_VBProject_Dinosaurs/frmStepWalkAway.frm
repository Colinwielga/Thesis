VERSION 5.00
Begin VB.Form frmStepWalkAway 
   Caption         =   "Walk Away"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   Picture         =   "frmStepWalkAway.frx":0000
   ScaleHeight     =   7350
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack10 
      BackColor       =   &H8000000E&
      Caption         =   "Back to the Main Page!"
      Height          =   855
      Left            =   4200
      MaskColor       =   &H8000000E&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblWalkAway 
      BackColor       =   &H8000000E&
      Caption         =   $"frmStepWalkAway.frx":3299F
      Height          =   1095
      Left            =   3360
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "frmStepWalkAway"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack10_Click()
    'This button will take the user back to the loading page.
    frmStepWalkAway.Visible = False
    frmLoad.Visible = True
End Sub

