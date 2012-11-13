VERSION 5.00
Begin VB.Form frmPicturesSarah 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Sarah"
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
      Left            =   15840
      TabIndex        =   3
      Top             =   12240
      Width           =   2175
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
      Left            =   15840
      TabIndex        =   2
      Top             =   9720
      Width           =   2175
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
      Height          =   2295
      Left            =   15840
      TabIndex        =   1
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Label lblSarah 
      Alignment       =   2  'Center
      Caption         =   "Me!"
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
      Left            =   9720
      TabIndex        =   0
      Top             =   12120
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   10095
      Left            =   6720
      Picture         =   "PicturesSarah.frx":0000
      Top             =   1680
      Width           =   7695
   End
End
Attribute VB_Name = "frmPicturesSarah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Address Book
'frmPicturesSarah(PicturesSarah.frm)
'Sarah Keating
'10-27-03
'Purpose: This form shows a picture of me.

Private Sub cmdNext_Click()
frmPicturesJessi.Show
frmPicturesSarah.Hide
' Allows the User to go to the next photo
End Sub

Private Sub cmdPrevious_Click()
frmPicturesSarah.Hide
frmPicturesAbbySuzyBree.Show
' Allows the user to go to the previous photo
End Sub

Private Sub cmdQuit_Click()
    End
    ' Allows the user to quit the program
End Sub
