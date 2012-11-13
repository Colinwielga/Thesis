VERSION 5.00
Begin VB.Form frmWar 
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Capitulate"
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
      Left            =   6240
      TabIndex        =   2
      Top             =   6000
      Width           =   2415
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Forward!"
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
      Left            =   6240
      TabIndex        =   1
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   $"frmElders1.frx":0000
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   5160
      Width           =   6135
   End
   Begin VB.Image Image1 
      Height          =   6750
      Left            =   -120
      Picture         =   "frmElders1.frx":0133
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmWar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form merely gives the user game feed back, and the option to quit or advance
'the game to a new form

Private Sub cmdNext_Click()
frmWar.Hide
frmArmy1.Show
End Sub
Private Sub cmdQuit_Click()
End
End Sub
