VERSION 5.00
Begin VB.Form frmStaring 
   Caption         =   "Embarrassment"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   Picture         =   "frmStaring.frx":0000
   ScaleHeight     =   4890
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack9 
      BackColor       =   &H8000000E&
      Caption         =   "Back to main page"
      Height          =   855
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblScream 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmStaring.frx":3D02
      Height          =   1095
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmStaring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack9_Click()
    'Nagivates the user back to the main page
    frmStaring.Visible = False
    frmLoad.Visible = True
End Sub
