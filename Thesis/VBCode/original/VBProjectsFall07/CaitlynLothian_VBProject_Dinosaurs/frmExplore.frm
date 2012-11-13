VERSION 5.00
Begin VB.Form frmExplore 
   BackColor       =   &H0080FFFF&
   Caption         =   "Explore the world"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Continue"
      Height          =   1095
      Left            =   1200
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lblExplore 
      BackColor       =   &H0080FFFF&
      Caption         =   "You are a brave soul! You venture forth into the unknown, ignorant of what you may find. "
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
End
Attribute VB_Name = "frmExplore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdContinue_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmExplore.Visible = False
    frmWaterHole.Visible = True
    
    'Message: find shoreline
    MsgBox "Good heavens " & YourName & ", you have found a small still pond that looks like it may be used as a watering hole. How thrilling!", , "Pond"
End Sub
