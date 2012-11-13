VERSION 5.00
Begin VB.Form frmIntro 
   Caption         =   "Ted Williams Introduction"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   Picture         =   "frmIntro.frx":0000
   ScaleHeight     =   7890
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   3855
   End
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H000000FF&
      Caption         =   " Learn about Ted Williams"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   3975
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnter_Click()
frmTedmenu.Show
frmIntro.Hide
End Sub

Private Sub cmdquit_Click()
End
End Sub
