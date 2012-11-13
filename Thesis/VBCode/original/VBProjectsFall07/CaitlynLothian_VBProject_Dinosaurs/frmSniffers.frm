VERSION 5.00
Begin VB.Form frmSniffers 
   BackColor       =   &H00404040&
   Caption         =   "Sniff sniff"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTurnAround 
      Caption         =   "Turn around, slowly and carefully."
      Height          =   735
      Left            =   2520
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdDive 
      BackColor       =   &H00404040&
      Caption         =   "Dive!"
      Height          =   735
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblDinoBreath 
      BackColor       =   &H00404040&
      Caption         =   $"frmSniffers.frx":0000
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmSniffers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDive_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmSniffers.Visible = False
    frmFishTeeth.Visible = True
    
    'Message about what happens
    MsgBox "You decide to jump in. The pond is much deeper than you expected, but before you get to dive much deeper, a set of very large, very sharp teeth enter your periphrial.", , "Fish?"
End Sub

Private Sub cmdTurnAround_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmSniffers.Visible = False
    frmDuckBill.Visible = True
    
    'Message box about first dinosaur encounter
    MsgBox "You turn around slowly and find yourself face to face with a gentle-looking duck-billed dinosaur. HOLY GOODNESS THAT'S A DINOSAUR!!!", , "Dinosaur!"
End Sub
