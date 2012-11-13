VERSION 5.00
Begin VB.Form frmPicturesNicole 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Nicole"
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
      Left            =   13440
      TabIndex        =   3
      Top             =   11880
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
      Height          =   1935
      Left            =   13440
      TabIndex        =   2
      Top             =   9480
      Width           =   2535
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
      Height          =   2175
      Left            =   13440
      TabIndex        =   1
      Top             =   6840
      Width           =   2535
   End
   Begin VB.Label lblNicole 
      Alignment       =   2  'Center
      Caption         =   "Nicole"
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
      Left            =   7560
      TabIndex        =   0
      Top             =   10800
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   5760
      Picture         =   "PicturesNicole.frx":0000
      Top             =   1560
      Width           =   6750
   End
End
Attribute VB_Name = "frmPicturesNicole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Address Book
'frmPicturesNicole(PicturesNicole.frm)
'Sarah Keating
'10-25-03
'Purpose: This form shows a picture of Nicole, my other roommate.

Private Sub cmdNext_Click()
frmPicturesAbbySuzyBree.Show
frmPicturesNicole.Hide
' Allows the User to go to the next photo
End Sub

Private Sub cmdPrevious_Click()
frmPicturesNicole.Hide
frmPicturesMara.Show
' Allows the user to go to the previous photo
End Sub

Private Sub cmdQuit_Click()
    End
    ' Allows the user to quit the program
End Sub
