VERSION 5.00
Begin VB.Form QuitShep 
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Text            =   "Where u think ur goinn?"
      Top             =   5040
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Click anywhere to exit."
      Top             =   0
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   8100
      Left            =   -360
      Picture         =   "QuitShep.frx":0000
      Top             =   -360
      Width           =   10800
   End
End
Attribute VB_Name = "QuitShep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Image1_Click()
End ' ends game
End Sub

