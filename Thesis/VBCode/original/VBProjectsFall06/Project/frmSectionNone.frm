VERSION 5.00
Begin VB.Form frmSectionNone 
   BackColor       =   &H00000080&
   Caption         =   "Not Available"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   4080
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   600
      Picture         =   "frmSectionNone.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   4515
      TabIndex        =   1
      Top             =   960
      Width           =   4575
   End
   Begin VB.Label lblSection23 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "No picture available"
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
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmSectionNone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmSectionNone
'Cole and John
'10/30/06
'Objective: The objective of this form is to tell the user they have chosen an
'invalid section number or a section number with no photo.  Thus, no picture exists.

Option Explicit

Private Sub cmdBack_Click()
    frmSectionNone.Hide     'see frmSection1 for comments
    frmSectionView.Show
End Sub
