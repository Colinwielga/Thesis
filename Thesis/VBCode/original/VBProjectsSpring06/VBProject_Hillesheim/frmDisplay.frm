VERSION 5.00
Begin VB.Form frmDisplay 
   Caption         =   "Leader Display"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   10590
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDisplay 
      Height          =   2415
      Left            =   4320
      ScaleHeight     =   2355
      ScaleWidth      =   3075
      TabIndex        =   9
      Top             =   840
      Width           =   3135
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command1"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command1"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "nag"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "yama"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    objDisplay.LoadPicture (Yamamoto)
    
End Sub

