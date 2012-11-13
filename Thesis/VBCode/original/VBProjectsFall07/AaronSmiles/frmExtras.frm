VERSION 5.00
Begin VB.Form frmExtras 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdAiDU 
      Caption         =   "Fight Test"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   3375
   End
   Begin VB.CommandButton cmdMArc 
      Caption         =   "Major Arcana"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   3375
   End
   Begin VB.CommandButton cmdFTest 
      Caption         =   "Ai Du"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmextras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A short playlist of random songs
Private Sub cmdAiDU_Click()
Shell ("explorer.exe M:\CS130\Try 3\01-Fight Test.mp3")
End Sub

Private Sub cmdexit_Click()
frmextras.Hide
End Sub

Private Sub cmdFTest_Click()
Shell ("explorer.exe M:\CS130\Try 3\09 ai du.mp3")
End Sub

Private Sub cmdMArc_Click()
Shell ("explorer.exe M:\CS130\Try 3\01.major arcana.mp3")
End Sub


