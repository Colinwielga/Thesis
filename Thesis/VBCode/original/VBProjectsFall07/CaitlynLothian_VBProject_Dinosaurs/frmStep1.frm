VERSION 5.00
Begin VB.Form frmStep1 
   BackColor       =   &H00FF0000&
   Caption         =   "The Introduction"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   Picture         =   "frmStep1.frx":0000
   ScaleHeight     =   3795
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWalkAway 
      Caption         =   "Walk Away"
      Height          =   855
      Left            =   6000
      TabIndex        =   2
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton cmdHelpHim 
      Caption         =   "Help the Old Man"
      Height          =   855
      Left            =   3120
      TabIndex        =   1
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label lblIntroductionTimeMachine 
      BackColor       =   &H00FF0000&
      Caption         =   $"frmStep1.frx":3016
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "frmStep1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHelpHim_Click()
    'Hides the form and reveals a new form
    frmStep1.Visible = False
    frmStepHelpHim.Visible = True
    
End Sub


Private Sub cmdWalkAway_Click()
    'Hides the form to reveal a new one
    frmStep1.Visible = False
    frmStepWalkAway.Visible = True
    
End Sub
