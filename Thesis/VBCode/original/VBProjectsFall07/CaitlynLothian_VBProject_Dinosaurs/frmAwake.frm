VERSION 5.00
Begin VB.Form frmAwake 
   Caption         =   "Awake"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   Picture         =   "frmAwake.frx":0000
   ScaleHeight     =   5220
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack1 
      BackColor       =   &H8000000E&
      Caption         =   "Back to Main Page"
      Height          =   735
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label lblAwake 
      BackColor       =   &H8000000E&
      Caption         =   $"frmAwake.frx":2AE14
      Height          =   1335
      Left            =   240
      TabIndex        =   1
      Top             =   2880
      Width           =   4215
   End
End
Attribute VB_Name = "frmAwake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack1_Click()
    'Navigates the user back to the main page
    frmAwake.Visible = False
    frmLoad.Visible = True
End Sub
