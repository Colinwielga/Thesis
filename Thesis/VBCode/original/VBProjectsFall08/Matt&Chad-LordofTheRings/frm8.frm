VERSION 5.00
Begin VB.Form frm8 
   Caption         =   "Form1"
   ClientHeight    =   6525
   ClientLeft      =   2220
   ClientTop       =   1710
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   Picture         =   "frm8.frx":0000
   ScaleHeight     =   6525
   ScaleWidth      =   6510
   Begin VB.CommandButton Command4 
      Caption         =   "End Your Journey"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Go Back"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Try to Escape"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stay And Fight"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frm8.frx":9179
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frm8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MsgBox ("You kill their leader Lurtz and wipeout the rest of them and although you are exhausted you feel very confident about the abilities of your fellowship.")
    frm8.Hide
    frmA.Show
End Sub

Private Sub Command2_Click()
    frm8.Hide
    frm9.Show
End Sub

Private Sub Command3_Click()
    frm8.Hide
    frm7.Show
End Sub

Private Sub Command4_Click()
    End
End Sub
