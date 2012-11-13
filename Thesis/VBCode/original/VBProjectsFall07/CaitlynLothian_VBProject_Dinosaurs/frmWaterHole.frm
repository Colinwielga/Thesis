VERSION 5.00
Begin VB.Form frmWaterHole 
   Caption         =   "Water Water Everywhere..."
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   Picture         =   "frmWaterHole.frx":0000
   ScaleHeight     =   5070
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNoDrink 
      BackColor       =   &H00808000&
      Caption         =   "Eww...No Way."
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdDrink 
      BackColor       =   &H00808000&
      Caption         =   "Chug it! Chug it!"
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblWaterHole 
      BackColor       =   &H00808000&
      Caption         =   $"frmWaterHole.frx":8C0F
      Height          =   1335
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmWaterHole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDrink_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmWaterHole.Visible = False
    frmSniffers.Visible = True
    
    'Message box for drink
    MsgBox "You bend down to drink the water. It's cool and refreshing, and you take comfort in the knowledge that you have found a steady supply of water. But just as you feel safe, you feel hot, dank breath on the back of your neck. Oh noes!", , "Surprise!"
End Sub

Private Sub cmdNoDrink_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmWaterHole.Visible = False
    frmWalking.Visible = True
End Sub
