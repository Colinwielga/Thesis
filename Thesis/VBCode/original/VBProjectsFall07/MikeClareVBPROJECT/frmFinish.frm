VERSION 5.00
Begin VB.Form frmFinish 
   Caption         =   "Form1"
   ClientHeight    =   9750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   Picture         =   "frmFinish.frx":0000
   ScaleHeight     =   9750
   ScaleWidth      =   13065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit game?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5520
      TabIndex        =   2
      Top             =   7440
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You fought your way to safety and beat the aliens for now!  Good work!"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1935
      Left            =   3480
      TabIndex        =   1
      Top             =   2760
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "YOU WIN!"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3015
      Left            =   3840
      TabIndex        =   0
      Top             =   1080
      Width           =   7575
   End
End
Attribute VB_Name = "frmFinish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdQuit_Click()
    End         'end the game
End Sub
