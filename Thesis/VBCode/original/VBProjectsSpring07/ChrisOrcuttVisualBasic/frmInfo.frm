VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H00000000&
   Caption         =   "Game Information"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConsoles 
      Caption         =   "Console Guide"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   2
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   3810
      Left            =   240
      Picture         =   "frmInfo.frx":0000
      Top             =   840
      Width           =   6660
   End
   Begin VB.Label lblReadUp 
      BackColor       =   &H00000000&
      Caption         =   "Questions about a Console? You'll find answers here!"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   10215
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConsoles_Click()
    frmInfo.Hide
    frmConsoleInfo.Show
End Sub
Private Sub cmdReturn_Click()
    frmInfo.Hide
    frmSelectWant.Show
End Sub

