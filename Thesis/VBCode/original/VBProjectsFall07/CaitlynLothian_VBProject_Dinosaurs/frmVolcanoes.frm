VERSION 5.00
Begin VB.Form frmVolcanoes 
   Caption         =   "HOT LAVA!!!"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   Picture         =   "frmVolcanoes.frx":0000
   ScaleHeight     =   5820
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGreen 
      BackColor       =   &H00008000&
      Height          =   855
      Left            =   2040
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdRed 
      BackColor       =   &H000000FF&
      Height          =   855
      Left            =   120
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblVolcanoes 
      BackColor       =   &H0080FFFF&
      Caption         =   $"frmVolcanoes.frx":54AA
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmVolcanoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGreen_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmVolcanoes.Visible = False
    frmExplosion.Visible = True
    
    'Message about the time machine's actions
    MsgBox "You pressed the green button. The machine whirs and cranks and jerks, and the volcanoes disappear. Everything stops.", , "Time Machine"
End Sub

Private Sub cmdRed_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmVolcanoes.Visible = False
    frmTimeMachine.Visible = True
    
    'Message about the time machine's actions
    MsgBox "Your logic in choosing the red is that the green takes you back in time. Red must mean forward.", , "Time Machine"
    MsgBox "The machine whirs and cranks and jerks, and the volcanoes disappear. When it stops, you realize that you're back in the jungle with the dinoaurs. Well, try again...", , "Time Machine"

End Sub
