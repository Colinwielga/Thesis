VERSION 5.00
Begin VB.Form frmSection1 
   BackColor       =   &H00000080&
   Caption         =   "Section 1"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   6690
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
      Picture         =   "frmSection1.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label lblSection1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "Section 1"
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
      TabIndex        =   1
      Top             =   240
      Width           =   1965
   End
End
Attribute VB_Name = "frmSection1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmSection1
'Cole and John
'10/30/06
'Objective: The objective of this form is to show the user what a view from this
'section looks like.

Option Explicit

Private Sub cmdBack_Click()
    frmSection1.Hide        'shows section1 form
    frmSectionView.Show
End Sub
