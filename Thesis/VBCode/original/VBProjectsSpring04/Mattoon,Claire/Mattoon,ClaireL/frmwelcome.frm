VERSION 5.00
Begin VB.Form frmwelcome 
   BackColor       =   &H0080C0FF&
   Caption         =   "Welcome"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSAT 
      BackColor       =   &H008080FF&
      Caption         =   "To predict college GPA based on a SAT score, please click here"
      Height          =   1215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdACT 
      BackColor       =   &H00FF8080&
      Caption         =   "To predict college GPA based on an ACT score, please click here"
      Height          =   1215
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Claire L. Mattoon"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblwelcome 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   $"frmwelcome.frx":0000
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmwelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdACT_Click()
frmACT.Show
End Sub

Private Sub cmdSAT_Click()
frmSAT.Show
End Sub
