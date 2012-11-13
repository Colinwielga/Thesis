VERSION 5.00
Begin VB.Form frmStepHelpHim 
   BackColor       =   &H00404040&
   Caption         =   "You help. You are a hero."
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   Picture         =   "frmStepHelpHim.frx":0000
   ScaleHeight     =   4860
   ScaleWidth      =   4185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   2280
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdForward 
      BackColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label lblTimeMachine 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmStepHelpHim.frx":22EB
      ForeColor       =   &H8000000E&
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmStepHelpHim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmStepHelpHim.Visible = False
    frmNewWorld.Visible = True
    
    'Displays a message in a message box about the arrival of the time machine
    MsgBox "Whoa, " & YourName & "! The machine has stopped, and you're in one piece! Hoozah!", , "Alive!"
    
    
End Sub

Private Sub cmdForward_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmStepHelpHim.Visible = False
    frmDeath.Visible = True
End Sub

