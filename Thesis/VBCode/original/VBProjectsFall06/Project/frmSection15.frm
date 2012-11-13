VERSION 5.00
Begin VB.Form frmSection15 
   BackColor       =   &H00000080&
   Caption         =   "Section 15"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   4080
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   840
      Picture         =   "frmSection15.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label lblSection15 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "Section 15"
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
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   2235
   End
End
Attribute VB_Name = "frmSection15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmSection15
'Cole and John
'10/30/06
'Objective: The objective of this form is to show the user what a view from this
'section looks like.

Option Explicit

Private Sub cmdBack_Click()
    frmSection15.Hide       'see frmSection1 for comments
    frmSectionView.Show
End Sub
