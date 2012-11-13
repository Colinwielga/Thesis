VERSION 5.00
Begin VB.Form frmFishTeeth 
   BackColor       =   &H8000000E&
   Caption         =   "CHOMP!"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   Picture         =   "frmFishTeeth.frx":0000
   ScaleHeight     =   6315
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack5 
      BackColor       =   &H00808000&
      Caption         =   "Back To Main Page"
      Height          =   855
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label lblFishTeeth 
      BackColor       =   &H8000000E&
      Caption         =   $"frmFishTeeth.frx":2D1E
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   3480
      Width           =   4095
   End
End
Attribute VB_Name = "frmFishTeeth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBack5_Click()
    'This button brings the user back to the loading page
    frmFishTeeth.Visible = False
    frmLoad.Visible = True
End Sub
