VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00808000&
   Caption         =   "Form3"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form3"
   ScaleHeight     =   7290
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   6960
      TabIndex        =   3
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton cmdhighscore 
      Caption         =   "View High Scores"
      Height          =   1095
      Left            =   4440
      TabIndex        =   2
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdplaygame 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Start a Game of Memory!"
      Default         =   -1  'True
      Height          =   1095
      Left            =   3120
      MaskColor       =   &H80000010&
      TabIndex        =   0
      Top             =   4800
      Width           =   4695
   End
   Begin VB.Label lblmenu 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Main Menu"
      Height          =   855
      Left            =   2040
      TabIndex        =   1
      Top             =   480
      Width           =   6615
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdquit_Click()
    End
    
End Sub
