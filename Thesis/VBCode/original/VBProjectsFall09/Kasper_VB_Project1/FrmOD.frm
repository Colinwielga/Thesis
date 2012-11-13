VERSION 5.00
Begin VB.Form FrmOD 
   Caption         =   "Offense Defense"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMain 
      Caption         =   "Return to main menu"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdDefense 
      Caption         =   "Defense"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdOffense 
      Caption         =   "Offense"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   4725
      Left            =   0
      Picture         =   "FrmOD.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6465
   End
End
Attribute VB_Name = "FrmOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Author: Brandon Kasper
'Written 10/19/2009
'this form lets the user decide to go to offense or deffense

Private Sub cmddefense_Click()
    FrmOD.Hide 'hides form from user
    frmDefense.Show 'shows form for user
End Sub

Private Sub cmdMain_Click()
    FrmOD.Hide 'hides OD page to user
    frmWelcome.Show 'shows Welcome page from user
End Sub

Private Sub cmdOffense_Click()
    FrmOD.Hide 'hides Welcome page from user
    frmStatsO.Show 'shows teams page to user
End Sub
