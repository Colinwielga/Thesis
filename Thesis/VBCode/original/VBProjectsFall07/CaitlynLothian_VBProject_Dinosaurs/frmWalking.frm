VERSION 5.00
Begin VB.Form frmWalking 
   BackColor       =   &H00004000&
   Caption         =   "Shuffle Shuffle Shuffle"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdForest 
      Caption         =   "Into the forest!"
      Height          =   855
      Left            =   2520
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmWalking.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdShoreline 
      Caption         =   "Around the shore!"
      Height          =   855
      Left            =   360
      Picture         =   "frmWalking.frx":14508
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblWalking 
      BackColor       =   &H00004000&
      Caption         =   $"frmWalking.frx":25B98
      ForeColor       =   &H8000000E&
      Height          =   1575
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmWalking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdForest_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmWalking.Visible = False
    frmMosquitoes.Visible = True
    
    'Message about presence of bugs
    MsgBox YourName & ", you head into the jungle. You walk through, looking at all of the exotic plants and realizing that your friends majoring in biology would love such an opprotunity. Yet, you hear a very loud buzzing, getting closer and closer.", , "Buzz"
End Sub

Private Sub cmdShoreline_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmWalking.Visible = False
    frmSqueaks.Visible = True
    
    'Message about walking
    MsgBox "Sure, the pond isn't the size of the lake, but it's large enough for you to explore.", , "Walk"
End Sub
