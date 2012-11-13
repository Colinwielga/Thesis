VERSION 5.00
Begin VB.Form frm10Die 
   Caption         =   "Form1"
   ClientHeight    =   7125
   ClientLeft      =   2220
   ClientTop       =   1710
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   Picture         =   "frm10Die.frx":0000
   ScaleHeight     =   7125
   ScaleWidth      =   9465
   Begin VB.CommandButton Command2 
      Caption         =   "End Your Journey"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go Back"
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   $"frm10Die.frx":8BB1
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "frm10Die"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frm10Die.Hide
    frm10.Show
End Sub

Private Sub Command2_Click()
    End
End Sub
