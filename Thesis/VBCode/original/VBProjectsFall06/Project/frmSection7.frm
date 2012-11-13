VERSION 5.00
Begin VB.Form frmSection7 
   BackColor       =   &H00000080&
   Caption         =   "Section 7"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   4320
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   960
      Picture         =   "frmSection7.frx":0000
      ScaleHeight     =   3075
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label lblSection7 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "Section 7"
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
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   1995
   End
End
Attribute VB_Name = "frmSection7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmSection7
'Cole and John
'10/30/06
'Objective: The objective of this form is to show the user what a view from this
'section looks like.

Option Explicit

Private Sub cmdBack_Click()
frmSection7.Hide            'see frmSection1 for comments
frmSectionView.Show
End Sub
