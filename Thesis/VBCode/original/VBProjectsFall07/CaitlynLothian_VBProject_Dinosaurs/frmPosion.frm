VERSION 5.00
Begin VB.Form frmPosion 
   BackColor       =   &H000000FF&
   Caption         =   "POSIONED!!!"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   Picture         =   "frmPosion.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack8 
      BackColor       =   &H8000000E&
      Caption         =   "Back to the Main Page"
      Height          =   975
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lblPosion 
      BackColor       =   &H8000000E&
      Caption         =   $"frmPosion.frx":2842
      Height          =   1695
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmPosion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack8_Click()
    'This button will take the user back to the loading page
    frmPosion.Visible = False
    frmLoad.Visible = True
End Sub

