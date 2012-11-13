VERSION 5.00
Begin VB.Form frmSection13 
   BackColor       =   &H00000080&
   Caption         =   "Section 13"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   4200
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   840
      Picture         =   "frmSection13.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label lblSection13 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "Section 13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   2235
   End
End
Attribute VB_Name = "frmSection13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmSection13
'Cole and John
'10/30/06
'Objective: The objective of this form is to show the user what a view from this
'section looks like.

Option Explicit

Private Sub cmdBack_Click()
    frmSection13.Hide       'see frmSection1 for comments
    frmSectionView.Show
End Sub
