VERSION 5.00
Begin VB.Form frmDeath 
   BackColor       =   &H80000012&
   Caption         =   "BIG BADDA BOOM"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   Picture         =   "frmDeath.frx":0000
   ScaleHeight     =   4770
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack4 
      BackColor       =   &H000080FF&
      Caption         =   "Back to Main Page"
      Height          =   735
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   $"frmDeath.frx":6AEB
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmDeath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack4_Click()
    'This button will take the user back to the loading page
    frmDeath.Visible = False
    frmLoad.Visible = True
End Sub
