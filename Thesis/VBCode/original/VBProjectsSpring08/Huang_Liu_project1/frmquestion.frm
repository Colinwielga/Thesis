VERSION 5.00
Begin VB.Form frmquestion 
   Caption         =   "Kitty's question"
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7755
   LinkTopic       =   "Form2"
   Picture         =   "frmquestion.frx":0000
   ScaleHeight     =   2430
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdstart 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Start Your Challenge!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   600
      MaskColor       =   &H00000000&
      Picture         =   "frmquestion.frx":1F99
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back to the main page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmquestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdback_Click()
frmmain.Visible = True
frmquestion.Visible = False
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdstart_Click()
frmquestion.Visible = False
frmquestion1.Visible = True
End Sub
