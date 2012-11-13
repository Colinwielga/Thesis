VERSION 5.00
Begin VB.Form frmLonevictor 
   Caption         =   "Lone Victor"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   Picture         =   "frmLonevictor.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000040C0&
      Caption         =   "Exit Program"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9960
      Width           =   2295
   End
End
Attribute VB_Name = "frmLonevictor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form presents the user with the final outcome of his decisions
'in the game
'it also gives him a command button allowing his exiting the program
Private Sub cmdQuit_Click()
End
End Sub

