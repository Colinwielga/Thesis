VERSION 5.00
Begin VB.Form frm11 
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   4005
   ClientTop       =   1320
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   Picture         =   "frm11.frx":0000
   ScaleHeight     =   8550
   ScaleWidth      =   6270
   Begin VB.CommandButton Command3 
      Caption         =   "Go Back"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "End Your Journey"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continue"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frm11.frx":745D
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   5640
      Width           =   6015
   End
End
Attribute VB_Name = "frm11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frm11.Hide
    frmWIN.Show
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Command3_Click()
    frm11.Hide
    frm10.Show
End Sub
