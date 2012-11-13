VERSION 5.00
Begin VB.Form frmSTART 
   BackColor       =   &H80000006&
   Caption         =   "Form1"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9495
   FillColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGermany 
      BackColor       =   &H80000007&
      Caption         =   "Germany"
      Height          =   495
      Left            =   4680
      MaskColor       =   &H8000000D&
      TabIndex        =   5
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSpain 
      Caption         =   "Spain"
      Height          =   495
      Left            =   7800
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdItaly 
      Caption         =   "Italy"
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdEngland 
      Caption         =   "England"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
   End
   Begin VB.PictureBox picDisplay 
      FillColor       =   &H0000FFFF&
      Height          =   5055
      Left            =   1080
      Picture         =   "frmSTART.frx":0000
      ScaleHeight     =   4995
      ScaleWidth      =   7395
      TabIndex        =   1
      Top             =   240
      Width           =   7455
   End
   Begin VB.Label lblLocation 
      BackColor       =   &H80000007&
      Caption         =   "What country do you want to see a soccer game in?"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   5880
      Width           =   1935
   End
End
Attribute VB_Name = "frmSTART"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEngland_Click()
    frmSTART.Hide
    frmEngland.Show
End Sub

