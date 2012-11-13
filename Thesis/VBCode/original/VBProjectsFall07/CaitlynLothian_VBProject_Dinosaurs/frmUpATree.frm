VERSION 5.00
Begin VB.Form frmUpATree 
   Caption         =   "Climb Up That Tree!"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   Picture         =   "frmUpATree.frx":0000
   ScaleHeight     =   5265
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPass 
      BackColor       =   &H00C0FFFF&
      Caption         =   "I'll pass, thanks."
      Height          =   855
      Left            =   240
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton cmdEat 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Yum yum yum, food in my belly!"
      Height          =   855
      Left            =   240
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lblTree 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmUpATree.frx":57F5
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmUpATree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEat_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmUpATree.Visible = False
    frmPosion.Visible = True
    
    'Informs the user that eating the fruit was a poor decision
    MsgBox YourName & " eats the fruit quickly, but within minutes, as you are relaxing in your tree, your stomach does not feel well. This is not good.", , "Icky"
End Sub


Private Sub cmdPass_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmUpATree.Visible = False
    frmExplore.Visible = True
    
    'Transition Message
    MsgBox "You are so hungry that you would rather find some safe food to eat. Sitting in this tree isn't going to help, so you climb back down to find food.", , "Leave"
End Sub
