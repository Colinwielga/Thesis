VERSION 5.00
Begin VB.Form frmKenMoney 
   BackColor       =   &H00000000&
   Caption         =   "Money"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4425
   BeginProperty Font 
      Name            =   "Harlow Solid Italic"
      Size            =   20.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picName 
      BackColor       =   &H00C00000&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   960
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   2
      Top             =   4800
      Width           =   2055
   End
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Image Image2 
      Height          =   2535
      Left            =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Image imgKen 
      Height          =   2535
      Left            =   -120
      Picture         =   "frmMoney.frx":0000
      Top             =   0
      Width           =   4470
   End
   Begin VB.Image imgStand 
      Height          =   4110
      Left            =   0
      Picture         =   "frmMoney.frx":2566
      Top             =   2520
      Width           =   4470
   End
End
Attribute VB_Name = "frmKenMoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGameBoard_Click()
    
    'Shows and hides the forms
    frmKenMoney.Hide
    frmGameBoard.Show

End Sub
