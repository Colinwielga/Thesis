VERSION 5.00
Begin VB.Form frmSubmission 
   BackColor       =   &H00000000&
   Caption         =   "Submission"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   Picture         =   "frmSubmission.frx":0000
   ScaleHeight     =   7320
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   3255
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"frmSubmission.frx":D74A
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   5280
      Width           =   9015
   End
End
Attribute VB_Name = "frmSubmission"
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
