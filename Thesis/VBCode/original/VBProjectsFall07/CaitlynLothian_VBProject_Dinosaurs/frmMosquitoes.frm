VERSION 5.00
Begin VB.Form frmMosquitoes 
   BackColor       =   &H00404000&
   Caption         =   "GIANT BLOOD SUCKERS!!!!"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   Picture         =   "frmMosquitoes.frx":0000
   ScaleHeight     =   6690
   ScaleWidth      =   3810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTimeMachine 
      BackColor       =   &H00004080&
      Caption         =   "To the Time Machine!"
      Height          =   975
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label lblMosquitoes 
      BackColor       =   &H00404000&
      Caption         =   $"frmMosquitoes.frx":CEB3
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   4440
      Width           =   3855
   End
End
Attribute VB_Name = "frmMosquitoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdTimeMachine_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmMosquitoes.Visible = False
    frmTimeMachine.Visible = True
    
    'Message about reasoning for leaving
    MsgBox "The giant mosquitoes were the last straw. It's time to get out of here!", , "Mosquitoes"
End Sub
