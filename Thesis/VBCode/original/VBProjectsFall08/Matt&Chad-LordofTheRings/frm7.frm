VERSION 5.00
Begin VB.Form frm7 
   Caption         =   "Form1"
   ClientHeight    =   8355
   ClientLeft      =   3795
   ClientTop       =   1710
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   Picture         =   "frm7.frx":0000
   ScaleHeight     =   8355
   ScaleWidth      =   6060
   Begin VB.CommandButton Command2 
      Caption         =   "End Your Journey"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continue"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frm7.frx":6336
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   2400
      Width           =   5535
   End
End
Attribute VB_Name = "frm7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frm7.Hide
    frm8.Show
End Sub

Private Sub Command2_Click()
    End
End Sub
