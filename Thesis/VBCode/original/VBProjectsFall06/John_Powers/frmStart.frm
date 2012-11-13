VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Start"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start!"
      Height          =   1215
      Left            =   840
      TabIndex        =   2
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox txtName 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label lblStart 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Welcome!  Enter your name and click Start! to start the test!"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdStart_Click()
    MyName = txtName.Text
    frmWater.Show
    frmStart.Hide
End Sub
