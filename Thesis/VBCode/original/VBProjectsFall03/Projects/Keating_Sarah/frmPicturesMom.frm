VERSION 5.00
Begin VB.Form frmPicturesMom 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   14955
   ScaleWidth      =   19080
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
      Height          =   2055
      Left            =   13200
      TabIndex        =   3
      Top             =   11760
      Width           =   2535
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Click to see the Next Picture"
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
      Left            =   13200
      TabIndex        =   2
      Top             =   9120
      Width           =   2535
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Click to see the Previous Picture"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   13200
      TabIndex        =   1
      Top             =   6240
      Width           =   2535
   End
   Begin VB.Label lblMom 
      Alignment       =   2  'Center
      Caption         =   "Sarah And Her Mom"
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
      Left            =   6960
      TabIndex        =   0
      Top             =   12120
      Width           =   4335
   End
   Begin VB.Image imgMom 
      Height          =   9000
      Left            =   5640
      Picture         =   "frmPicturesMom.frx":0000
      Top             =   2640
      Width           =   6750
   End
End
Attribute VB_Name = "frmPicturesMom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Project Name: Address Book
'frmPicturesMom(PicturesMom.frm)
'Sarah Keating
'10-25-03
'Purpose: This form shows a picture of my Mom.

Option Explicit

Private Sub cmdNext_Click()
frmPicturesMom.Hide
frmPicturesMara.Show
' Allows the user to go to the next photo
End Sub

Private Sub cmdPrevious_Click()
frmPicturesMom.Hide
frmPicturesDad.Show
' Allows the user to go to the previous photo
End Sub

Private Sub cmdQuit_Click()
    End
    ' Allows the user to quit the program
End Sub
