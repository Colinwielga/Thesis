VERSION 5.00
Begin VB.Form frmBioTeacher 
   BackColor       =   &H00400000&
   Caption         =   "Wakey Wakey!"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmBioTeacher.frx":0000
   ScaleHeight     =   7170
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack2 
      Caption         =   "Back to the main page"
      Height          =   735
      Left            =   1440
      TabIndex        =   0
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label lblBioTeacher 
      BackColor       =   &H00400000&
      Caption         =   $"frmBioTeacher.frx":3ED3
      ForeColor       =   &H8000000E&
      Height          =   2055
      Left            =   0
      TabIndex        =   1
      Top             =   4560
      Width           =   4695
   End
End
Attribute VB_Name = "frmBioTeacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBack2_Click()
    'Brings the user back to the loading page
    frmBioTeacher.Visible = False
    frmLoad.Visible = True
End Sub
