VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main Form"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   5640
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Look at Existing Setlists"
      Height          =   1095
      Left            =   3000
      TabIndex        =   2
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create a setlist"
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Setlist Helper"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   5775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

frmMain.Hide

frmCreate.Show



End Sub

Private Sub Command2_Click()

frmMain.Hide

frmExist.Show

End Sub

Private Sub Command4_Click()
End
End Sub
