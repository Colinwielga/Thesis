VERSION 5.00
Begin VB.Form frmMsgresponse 
   Caption         =   "Form1"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   10965
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
      Left            =   7800
      TabIndex        =   2
      Top             =   7200
      Width           =   2415
   End
   Begin VB.CommandButton cmdOnward 
      Caption         =   "Onward!"
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
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lblInstructions 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMsgresponse.frx":0000
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   8640
      Left            =   0
      Picture         =   "frmMsgresponse.frx":0118
      Top             =   0
      Width           =   10995
   End
End
Attribute VB_Name = "frmMsgresponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form merely gives the user game feed back, and the option to quit or advance
'the game to a new form

Private Sub cmdOnward_Click()
frmMsgresponse.Hide
frmArmy1.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub
