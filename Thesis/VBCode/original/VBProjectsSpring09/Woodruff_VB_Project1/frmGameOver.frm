VERSION 5.00
Begin VB.Form frmGameOver 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H80000015&
      Caption         =   "Quit"
      Height          =   800
      Left            =   9960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7800
      Width           =   2500
   End
   Begin VB.CommandButton cmdRetry 
      BackColor       =   &H80000015&
      Caption         =   "Retry?"
      Height          =   800
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7800
      Width           =   2500
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "You Died"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   96
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3375
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   9255
   End
   Begin VB.Label lblGameOver 
      BackColor       =   &H80000012&
      Caption         =   "Game Over"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   96
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   5055
      Left            =   5880
      TabIndex        =   0
      Top             =   3720
      Width           =   7455
   End
End
Attribute VB_Name = "frmGameOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name:  Super Awesome Cave Adventure Game
'Form Name:  frmGameOver
'Author:  Peter Woodruff
'Date Written:  3-20-09
'Purpose:  This form is where the user goes when his character dies.
Option Explicit

Private Sub cmdQuit_Click()

    'End
    End
    
End Sub

Private Sub cmdRetry_Click()
    
    'Starts the game over
    frmGameOver.Visible = False
    frmTitle.Visible = True
    
End Sub

Private Sub Form_Load()

End Sub
