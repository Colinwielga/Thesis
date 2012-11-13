VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H000000FF&
   Caption         =   "Introduction"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H000000FF&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdSit2 
      Caption         =   "Situation 2"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdSit1 
      Caption         =   "Situation 1"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   $"frmIntro.frx":0000
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdSit1_Click()
    frmIntro.Hide
    frmSit1.Show
End Sub

Private Sub cmdSit2_Click()
    frmIntro.Hide
    frmSit2.Show
End Sub


