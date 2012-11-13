VERSION 5.00
Begin VB.Form frmTimeMachine 
   BackColor       =   &H00000000&
   Caption         =   "Time to go home"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGreen 
      BackColor       =   &H00008000&
      Height          =   855
      Left            =   2520
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdRed 
      BackColor       =   &H000000FF&
      Height          =   855
      Left            =   360
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblGoHome 
      BackColor       =   &H80000012&
      Caption         =   $"frmTimeMachine.frx":0000
      ForeColor       =   &H8000000E&
      Height          =   1695
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmTimeMachine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGreen_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmTimeMachine.Visible = False
    frmVolcanoes.Visible = True
    
    'Message about time machine's actions
    MsgBox "The machine whirs and cranks and jerks, and the jungle disappears. You remember that the green button sent you back, so you hope that the green button will also take you home.", , "Time Machine"
End Sub

Private Sub cmdRed_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmTimeMachine.Visible = False
    frmHome.Visible = True
    
    'Message about time machine's actions
    MsgBox "The machine whirs and cranks and jerks, and the jungle disappears. You hope this was the button to get you home.", , "Time Machine"
End Sub
