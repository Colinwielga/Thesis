VERSION 5.00
Begin VB.Form QitMtn 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11385
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
      Left            =   120
      TabIndex        =   1
      Text            =   "Kiss before you go?"
      Top             =   7680
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Click anywhere to quit."
      Top             =   0
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   8520
      Left            =   0
      Picture         =   "QitMtn.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11400
   End
End
Attribute VB_Name = "QitMtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Image1_Click()
End ' ends game
End Sub
