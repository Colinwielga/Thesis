VERSION 5.00
Begin VB.Form frmMountain 
   Caption         =   "Mountain!"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   Picture         =   "fmrMountain.frx":0000
   ScaleHeight     =   4455
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdKeepClimbing 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Keep on climbing"
      Height          =   735
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdCave 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Go into the cave"
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label lbMountainCave 
      BackColor       =   &H00C0E0FF&
      Caption         =   $"fmrMountain.frx":6295
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmMountain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCave_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmMountain.Visible = False
    frmCave.Visible = True
    
    'Message box about decision to enter cave
    MsgBox YourName & ", you enter the dark cave.", , "Cave"
End Sub

Private Sub cmdKeepClimbing_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmMountain.Visible = False
    frmMountainTop.Visible = True
    
    'Message box about rewards of the choice
    MsgBox YourName & ", it was a good thing you decided to keep climbing, because just as you begin your ascent, you see a pterodactyl fly out of the cave. You would have been it's lunch, and you're sure of it.", , "Good Thing You Continued"
End Sub
