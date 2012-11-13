VERSION 5.00
Begin VB.Form frmChasm 
   Caption         =   "Running"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   Picture         =   "frmChasm.frx":0000
   ScaleHeight     =   7530
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmNoScream 
      Caption         =   "No screams for you"
      Height          =   975
      Left            =   5400
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdScream 
      Caption         =   "You scream"
      Height          =   975
      Left            =   3240
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblRun 
      Caption         =   $"frmChasm.frx":F943
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "frmChasm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdScream_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmChasm.Visible = False
    frmStaring.Visible = True
End Sub

Private Sub cmNoScream_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmChasm.Visible = False
    frmAwake.Visible = True
End Sub
