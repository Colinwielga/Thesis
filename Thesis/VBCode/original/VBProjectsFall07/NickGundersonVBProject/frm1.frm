VERSION 5.00
Begin VB.Form frm1 
   BackColor       =   &H8000000D&
   Caption         =   "Welcome"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11235
   FillColor       =   &H000000FF&
   FillStyle       =   6  'Cross
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H8000000D&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5760
      TabIndex        =   2
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdProceed 
      BackColor       =   &H8000000D&
      Caption         =   "Go to the Match-Up"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   1
      Top             =   5520
      Width           =   1935
   End
   Begin VB.PictureBox picEnter 
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   3600
      Picture         =   "frm1.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this right here is your basic from just showing the project.  It allows you to end the project
'as well as start the project and proceed onto the next screen

Private Sub cmdProceed_Click()
    frm1.Hide
    frmMatchup.Show
    
End Sub

Private Sub cmdQuit_Click()
End
End Sub

