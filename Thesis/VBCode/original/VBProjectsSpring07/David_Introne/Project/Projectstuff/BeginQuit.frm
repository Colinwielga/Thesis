VERSION 5.00
Begin VB.Form BeginQuit 
   Caption         =   "Going already?"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   5955
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
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Text            =   "Smell you later!"
      Top             =   7680
      Width           =   2535
   End
   Begin VB.TextBox end 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Click anywhere to exit."
      Top             =   0
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   8040
      Left            =   0
      Picture         =   "BeginQuit.frx":0000
      Top             =   0
      Width           =   5985
   End
End
Attribute VB_Name = "BeginQuit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Image1_Click()
    End ' Quits Game
End Sub
