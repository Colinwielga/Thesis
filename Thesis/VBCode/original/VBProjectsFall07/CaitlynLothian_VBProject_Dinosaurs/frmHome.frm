VERSION 5.00
Begin VB.Form frmHome 
   BackColor       =   &H0080C0FF&
   Caption         =   "Back home"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   Picture         =   "frmHome.frx":0000
   ScaleHeight     =   6540
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack7 
      BackColor       =   &H8000000E&
      Caption         =   "Back to the main page"
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label lblLab 
      BackColor       =   &H8000000E&
      Caption         =   $"frmHome.frx":A0B0
      Height          =   2775
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack7_Click()
    'Brings the user back to the loading page
    frmHome.Visible = False
    frmLoad.Visible = True
End Sub
