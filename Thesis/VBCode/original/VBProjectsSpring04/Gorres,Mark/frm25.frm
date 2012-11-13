VERSION 5.00
Begin VB.Form LS 
   BackColor       =   &H80000008&
   Caption         =   "Lincoln LS"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   8040
      TabIndex        =   1
      Top             =   5760
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   5295
      Left            =   120
      Picture         =   "frm25.frx":0000
      ScaleHeight     =   5235
      ScaleWidth      =   8835
      TabIndex        =   0
      Top             =   120
      Width           =   8895
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Lincoln LS"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   5520
      Width           =   2055
   End
End
Attribute VB_Name = "LS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    LS.Hide
End Sub

