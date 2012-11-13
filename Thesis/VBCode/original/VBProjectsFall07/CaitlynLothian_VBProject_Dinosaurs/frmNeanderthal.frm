VERSION 5.00
Begin VB.Form frmNeanderthal 
   BackColor       =   &H00800080&
   Caption         =   "Poke poke poke"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTurn 
      Caption         =   "Turn, slowly"
      Height          =   735
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run Away!"
      Height          =   735
      Left            =   2400
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblNeanderthal 
      BackColor       =   &H00800080&
      Caption         =   $"frmNeanderthal.frx":0000
      ForeColor       =   &H8000000E&
      Height          =   1935
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmNeanderthal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRun_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmNeanderthal.Visible = False
    frmTimeMachine.Visible = True
End Sub

Private Sub cmdTurn_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmNeanderthal.Visible = False
    frmBioTeacher.Visible = True
End Sub
