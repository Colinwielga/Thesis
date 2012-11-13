VERSION 5.00
Begin VB.Form frmCave 
   Caption         =   "Cave"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   Picture         =   "frmCave.frx":0000
   ScaleHeight     =   5700
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Back to the main page"
      Height          =   855
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label lblPterodactyl 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmCave.frx":1730B
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   3855
   End
End
Attribute VB_Name = "frmCave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack3_Click()
    'Brings the user back to the loading page
    frmCave.Visible = False
    frmLoad.Visible = True
End Sub
