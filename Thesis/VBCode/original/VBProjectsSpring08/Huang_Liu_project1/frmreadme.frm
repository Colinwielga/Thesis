VERSION 5.00
Begin VB.Form frmreadme 
   Caption         =   "Read Me"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8835
   LinkTopic       =   "Form2"
   Picture         =   "frmreadme.frx":0000
   ScaleHeight     =   2700
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      Height          =   855
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back to the main page"
      Height          =   855
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Thanks to google. Sources are found on www.sanrio.com and www.wallcoo.com"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   8175
   End
   Begin VB.Label lblreadme 
      BackColor       =   &H00FFFFFF&
      Caption         =   "This Project is created by Tino Huang and Debby Liu."
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   8175
   End
End
Attribute VB_Name = "frmreadme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
frmmain.Visible = True
frmreadme.Visible = False
End Sub

Private Sub cmdquit_Click()
End
End Sub
