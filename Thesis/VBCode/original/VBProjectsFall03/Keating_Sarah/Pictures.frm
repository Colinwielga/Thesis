VERSION 5.00
Begin VB.Form frmPictures 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   6840
      ScaleHeight     =   2715
      ScaleWidth      =   2835
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.PictureBox picResultsPapaAndMe 
      Height          =   2775
      Left            =   3600
      ScaleHeight     =   2715
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.PictureBox picResultsMomandMe 
      Height          =   2775
      Left            =   360
      ScaleHeight     =   2715
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "frmPictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Address Book
'frmPictures(Pictures.frm)
'Sarah Keating
'10-25-03
'Purpose: This form shows multiple pictures of the people in this address book.

Option Explicit

Private Sub picResultsMomandMe_Click()
picResults.Picture = LoadPicture("M:\CS130\Projects\Keating_Sarah\Pictures_Mom")
End Sub
