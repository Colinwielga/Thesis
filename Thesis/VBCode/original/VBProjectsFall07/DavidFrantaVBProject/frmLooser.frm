VERSION 5.00
Begin VB.Form frmLooser 
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   4575
   ClientTop       =   3360
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   7725
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End Johnyopoly"
      Height          =   975
      Left            =   1440
      TabIndex        =   2
      Top             =   2280
      Width           =   4815
   End
   Begin VB.PictureBox picLooser 
      Height          =   735
      Left            =   2280
      ScaleHeight     =   675
      ScaleWidth      =   3075
      TabIndex        =   1
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "You Loose! Don't try a career in business!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   6375
   End
End
Attribute VB_Name = "frmLooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ends the game
Private Sub cmdEnd_Click()
End
End Sub

