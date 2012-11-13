VERSION 5.00
Begin VB.Form frmExplosion 
   Caption         =   "BIG BADDA BOOM!"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   Picture         =   "frmExplosion.frx":0000
   ScaleHeight     =   4455
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack5 
      BackColor       =   &H000080FF&
      Caption         =   "Back to the main page"
      Height          =   855
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label lblExplosion 
      BackColor       =   &H000080FF&
      Caption         =   $"frmExplosion.frx":6AEB
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmExplosion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack5_Click()
    'Brings the user back to the loading page
    frmExplosion.Visible = False
    frmLoad.Visible = True
End Sub
