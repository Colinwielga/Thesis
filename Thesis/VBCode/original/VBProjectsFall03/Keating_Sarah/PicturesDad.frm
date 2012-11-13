VERSION 5.00
Begin VB.Form frmPicturesDad 
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
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   15240
      TabIndex        =   2
      Top             =   12720
      Width           =   2535
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Click to go to the Next Picture"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   15240
      TabIndex        =   1
      Top             =   10200
      Width           =   2535
   End
   Begin VB.Label lblDad 
      Alignment       =   2  'Center
      Caption         =   "Sarah And Her Dad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      TabIndex        =   0
      Top             =   10800
      Width           =   4455
   End
   Begin VB.Image imgDad 
      Height          =   6750
      Left            =   5520
      Picture         =   "PicturesDad.frx":0000
      Top             =   3600
      Width           =   9000
   End
End
Attribute VB_Name = "frmPicturesDad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Address Book
'frmPicturesDad(PicturesDad.frm)
'Sarah Keating
'10-25-03
'Purpose: This form shows multiple pictures of the people in this address book.

Option Explicit




Private Sub cmdNext_Click()
frmPicturesMom.Show
frmPicturesDad.Hide
' Allows the user to go to the next picture/form
End Sub

Private Sub cmdQuit_Click()
    End
' Allows the user to quit the program
End Sub
