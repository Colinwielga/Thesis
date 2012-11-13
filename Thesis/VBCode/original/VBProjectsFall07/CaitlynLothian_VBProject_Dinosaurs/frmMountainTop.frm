VERSION 5.00
Begin VB.Form frmMountainTop 
   Caption         =   "Top o' the Mountain to ya!"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   Picture         =   "frmMountainTop.frx":0000
   ScaleHeight     =   6300
   ScaleWidth      =   4260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H00FFFF80&
      Caption         =   "Continue"
      Height          =   855
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label lblMountainTop 
      BackColor       =   &H00FFFF80&
      Caption         =   $"frmMountainTop.frx":AA32
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   3600
      Width           =   3855
   End
End
Attribute VB_Name = "frmMountainTop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdContinue_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmMountainTop.Visible = False
    frmNeanderthal.Visible = True
    
    'Message about next encounter
    MsgBox "You are clamly surveying the land, realizing that just earlier this morning, you were in a different place and a different time. Of course, you're soon interrupted...", , "Interruption"
End Sub
