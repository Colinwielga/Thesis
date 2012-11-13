VERSION 5.00
Begin VB.Form frmPicturesJessi 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Jessi"
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
      Left            =   16440
      TabIndex        =   3
      Top             =   12720
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
      Left            =   16440
      TabIndex        =   2
      Top             =   10320
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
      Left            =   16440
      TabIndex        =   1
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Label lblJessi 
      Alignment       =   2  'Center
      Caption         =   "Jessi"
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
      Left            =   9600
      TabIndex        =   0
      Top             =   11400
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   7980
      Left            =   5160
      Picture         =   "PicturesJessi.frx":0000
      Top             =   2880
      Width           =   10605
   End
End
Attribute VB_Name = "frmPicturesJessi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Address Book
'frmPicturesJessi(PicturesJessi.frm)
'Sarah Keating
'10-27-03
'Purpose: This form shows a picture of Jessi.

Private Sub cmdNext_Click()
frmPicturesKatieF.Show
frmPicturesJessi.Hide
' Allows the User to go to the next photo
End Sub

Private Sub cmdPrevious_Click()
frmPicturesJessi.Hide
frmPicturesSarah.Show
' Allows the user to go to the previous photo
End Sub

Private Sub cmdQuit_Click()
    End
    ' Allows the user to quit the program
End Sub
