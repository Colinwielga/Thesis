VERSION 5.00
Begin VB.Form frmBushMoney 
   BackColor       =   &H00000000&
   Caption         =   "Money"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Freestyle Script"
      Size            =   21.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGameBoard 
      BackColor       =   &H00C00000&
      Caption         =   "Game Board"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox picWinnings 
      BackColor       =   &H00C00000&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   600
      ScaleHeight     =   795
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   2640
      Width           =   2775
   End
   Begin VB.PictureBox picName 
      BackColor       =   &H00C00000&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   960
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Image imgStand 
      Height          =   4110
      Left            =   0
      Picture         =   "frmBushMoney.frx":0000
      Top             =   2160
      Width           =   4470
   End
   Begin VB.Image imgBush 
      Height          =   2265
      Left            =   120
      Picture         =   "frmBushMoney.frx":414F
      Top             =   0
      Width           =   4170
   End
End
Attribute VB_Name = "frmBushMoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This program will show the gameboard and hide the character

Private Sub cmdGameBoard_Click()
        
    'Shows and hides the forms
    frmBushMoney.Hide
    frmGameBoard.Show

End Sub
