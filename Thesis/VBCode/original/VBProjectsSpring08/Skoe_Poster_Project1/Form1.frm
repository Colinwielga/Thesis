VERSION 5.00
Begin VB.Form start 
   Caption         =   "Form1"
   ClientHeight    =   8835
   ClientLeft      =   660
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   10995
   Begin VB.CommandButton Command3 
      Caption         =   "Quit"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Use Character"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      Height          =   8175
      Left            =   2280
      ScaleHeight     =   8115
      ScaleWidth      =   8595
      TabIndex        =   1
      Top             =   0
      Width           =   8655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View Character"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   2415
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Char = InputBox("Please Enter Number")
picResults.Picture = LoadPicture(App.Path & "\" & Pics(Char))
MsgBox (Characters(Char) & " has been loaded.")

End Sub

Private Sub Command2_Click()
start.Hide
damage.Show
End Sub

