VERSION 5.00
Begin VB.Form frmAssassinationfailure 
   BackColor       =   &H00000040&
   Caption         =   "Failure"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   Picture         =   "frmAssassinationfailure.frx":0000
   ScaleHeight     =   7815
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00808080&
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
      Height          =   975
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H00808080&
      Caption         =   $"frmAssassinationfailure.frx":15BBD
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   1
      Top             =   5400
      Width           =   6255
   End
End
Attribute VB_Name = "frmAssassinationfailure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form simply gives the user feedback as to the ultimate outcome of his decisions
'througout the game, which is in this case failure and his death.

Private Sub cmdQuit_Click()
End
End Sub
