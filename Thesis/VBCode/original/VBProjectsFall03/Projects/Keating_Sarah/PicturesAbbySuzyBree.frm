VERSION 5.00
Begin VB.Form frmPicturesAbbySuzyBree 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Abby, Suzy, and Bree"
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
      Left            =   16080
      TabIndex        =   3
      Top             =   12720
      Width           =   2415
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
      Height          =   2055
      Left            =   16080
      TabIndex        =   2
      Top             =   10200
      Width           =   2415
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
      Height          =   2415
      Left            =   16080
      TabIndex        =   1
      Top             =   7320
      Width           =   2415
   End
   Begin VB.Label lblGirls 
      Alignment       =   2  'Center
      Caption         =   "Abby, Suzy, and Bree"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7680
      TabIndex        =   0
      Top             =   11400
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   7845
      Left            =   4560
      Picture         =   "PicturesAbbySuzyBree.frx":0000
      Top             =   3120
      Width           =   10860
   End
End
Attribute VB_Name = "frmPicturesAbbySuzyBree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Address Book
'frmPicturesAbbySuzyBree(PicturesAbbySuzyBree.frm)
'Sarah Keating
'10-25-03
'Purpose: This form shows a picture of Abby, Suzy, and Bree.

Private Sub cmdNext_Click()
frmPicturesSarah.Show
frmPicturesAbbySuzyBree.Hide
' Allows the User to go to the next photo
End Sub

Private Sub cmdPrevious_Click()
frmPicturesAbbySuzyBree.Hide
frmPicturesNicole.Show
' Allows the user to go to the previous photo
End Sub

Private Sub cmdQuit_Click()
    End
    ' Allows the user to quit the program
End Sub
