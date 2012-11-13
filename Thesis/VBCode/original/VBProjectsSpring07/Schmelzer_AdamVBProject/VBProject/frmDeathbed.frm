VERSION 5.00
Begin VB.Form frmDeathbed 
   Caption         =   "Defeat"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13860
   LinkTopic       =   "Form1"
   Picture         =   "frmDeathbed.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   13860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
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
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label lblInstructions2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDeathbed.frx":10FA4
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   7200
      TabIndex        =   1
      Top             =   5520
      Width           =   6135
   End
   Begin VB.Label lblInstructions 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDeathbed.frx":111D0
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   7200
      TabIndex        =   0
      Top             =   5520
      Width           =   6015
   End
End
Attribute VB_Name = "frmDeathbed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form presents the user with one of the forms that is a final outcome
'within this form one of two messages is conveyed to the user
'the one that is made visible is dependent upon whether or not the user
'delcared that he was a lover in the beginning of the game which would have
'changed the lover boolean variable to true from false

Private Sub cmdQuit_Click()
End
End Sub

Private Sub Form_Load()
If Lover = True Then
    frm1on1.Hide
    lblInstructions2.Visible = True
    lblInstructions.Visible = False
End If
If Lover = False Then
    frm1on1.Hide
    lblInstructions2.Visible = False
    lblInstructions.Visible = True
End If
End Sub
