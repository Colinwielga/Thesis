VERSION 5.00
Begin VB.Form frmNewWorld 
   BackColor       =   &H0000C000&
   Caption         =   "Where Am I?"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide me!"
      Height          =   735
      Left            =   3600
      TabIndex        =   3
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton cmdExplore 
      Caption         =   "Explore the new world!"
      Height          =   735
      Left            =   720
      TabIndex        =   2
      Top             =   5640
      Width           =   1935
   End
   Begin VB.PictureBox picJungle 
      Height          =   3855
      Left            =   0
      Picture         =   "frmNewWorld.frx":0000
      ScaleHeight     =   3795
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label lblJungleDecision 
      BackColor       =   &H0000C000&
      Caption         =   $"frmNewWorld.frx":14508
      Height          =   1455
      Left            =   480
      TabIndex        =   1
      Top             =   3960
      Width           =   5295
   End
End
Attribute VB_Name = "frmNewWorld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExplore_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmNewWorld.Visible = False
    frmExplore.Visible = True
End Sub

Private Sub cmdHide_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmNewWorld.Visible = False
    frmUpATree.Visible = True
End Sub
