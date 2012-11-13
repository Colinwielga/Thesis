VERSION 5.00
Begin VB.Form frmVictorfornow 
   BackColor       =   &H00000080&
   Caption         =   "Victory..."
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   Picture         =   "frmVictorfornow.frx":0000
   ScaleHeight     =   7530
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00000080&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   3375
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H00000080&
      Caption         =   "You have proved victorious in this instance, only time will prove your true worth. "
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   6480
      Width           =   5175
   End
End
Attribute VB_Name = "frmVictorfornow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form presents the user with the final outcome of his decisions
'in the game
'it also gives him a command button allowing his exiting the program
Private Sub cmdExit_Click()
End
End Sub
