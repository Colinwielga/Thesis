VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H0080C0FF&
   Caption         =   "Welcome To Bingo!!"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   5370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   3480
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   855
      Left            =   2160
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Label lblName 
      Caption         =   "Please Enter Your Name:"
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdStart_Click()
    UserName = txtName.Text
    txtWelcome.Text = ("Welcome to the game " & UserName)
    txtName.Visible = False
    lblName.Visible = False
End Sub

