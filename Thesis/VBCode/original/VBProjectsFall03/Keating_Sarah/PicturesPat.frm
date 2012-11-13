VERSION 5.00
Begin VB.Form frmPicturesPat 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Pat"
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
      Height          =   1815
      Left            =   16560
      TabIndex        =   3
      Top             =   12960
      Width           =   1935
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
      Height          =   1935
      Left            =   16560
      TabIndex        =   2
      Top             =   10560
      Width           =   1935
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Click to go to the Previous Picture"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   16560
      TabIndex        =   1
      Top             =   8040
      Width           =   1935
   End
   Begin VB.Label lblPat 
      Alignment       =   2  'Center
      Caption         =   "Pat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9000
      TabIndex        =   0
      Top             =   12000
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   8940
      Left            =   3960
      Picture         =   "PicturesPat.frx":0000
      Top             =   2640
      Width           =   11925
   End
End
Attribute VB_Name = "frmPicturesPat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Address Book
'frmPicturesPat(PicturesPat.frm)
'Sarah Keating
'10-27-03
'Purpose: This form shows a picture of Pat.

Private Sub cmdNext_Click()
frmPicturesJason.Show
frmPicturesPat.Hide
' Allows the User to go to the next photo
End Sub

Private Sub cmdPrevious_Click()
frmPicturesPat.Hide
frmPicturesKatieF.Show
' Allows the user to go to the previous photo
End Sub

Private Sub cmdQuit_Click()
    End
    ' Allows the user to quit the program
End Sub
