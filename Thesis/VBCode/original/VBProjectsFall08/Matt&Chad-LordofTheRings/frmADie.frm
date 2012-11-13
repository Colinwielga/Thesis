VERSION 5.00
Begin VB.Form frmADie 
   ClientHeight    =   7125
   ClientLeft      =   2625
   ClientTop       =   1905
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   Picture         =   "frmADie.frx":0000
   ScaleHeight     =   7125
   ScaleWidth      =   9465
   Begin VB.CommandButton Command2 
      Caption         =   "End Your Journey"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go Back"
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   $"frmADie.frx":8BB1
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmaDie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmaDie.Hide
    frm8.Show
End Sub

Private Sub Command2_Click()
    End
End Sub
