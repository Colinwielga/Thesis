VERSION 5.00
Begin VB.Form frmGlory 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Glory"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   Picture         =   "frmGlory.frx":0000
   ScaleHeight     =   8340
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
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
      Left            =   2040
      TabIndex        =   1
      Top             =   7680
      Width           =   3375
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"frmGlory.frx":1169B
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   5640
      Width           =   7455
   End
End
Attribute VB_Name = "frmGlory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form presents the user with the best and final outcome of his decisions
'in the game
'it also gives him a command button allowing his exiting the program

Private Sub cmdExit_Click()
End
End Sub
