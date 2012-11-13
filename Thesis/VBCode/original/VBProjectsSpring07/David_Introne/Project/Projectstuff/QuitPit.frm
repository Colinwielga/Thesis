VERSION 5.00
Begin VB.Form QuitPit 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   10815
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Text            =   "Your Leaving Me? Wimper..."
      Top             =   6720
      Width           =   4215
   End
   Begin VB.TextBox txtlbl 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Click anywhere to quit."
      Top             =   0
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   -480
      Picture         =   "QuitPit.frx":0000
      Top             =   -720
      Width           =   12000
   End
End
Attribute VB_Name = "QuitPit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Image1_Click()
End ' ends game
End Sub
