VERSION 5.00
Begin VB.Form frmSection3 
   BackColor       =   &H00000080&
   Caption         =   "Section 3"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   Picture         =   "frmSection3.frx":0000
   ScaleHeight     =   5025
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   4320
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   1080
      Picture         =   "frmSection3.frx":0342
      ScaleHeight     =   2955
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Label lblSection3 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "Section 3"
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
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   1995
   End
End
Attribute VB_Name = "frmSection3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmSection3
'Cole and John
'10/30/06
'Objective: The objective of this form is to show the user what a view from this
'section looks like.

Option Explicit

Private Sub cmdBack_Click()
    frmSection3.Hide        'see frmSection1 for comments
    frmSectionView.Show
End Sub
