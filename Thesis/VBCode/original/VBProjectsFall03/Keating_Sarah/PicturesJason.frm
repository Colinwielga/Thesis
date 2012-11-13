VERSION 5.00
Begin VB.Form frmPicturesJason 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Jason"
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
      Height          =   1935
      Left            =   16680
      TabIndex        =   3
      Top             =   11640
      Width           =   2055
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
      Left            =   16680
      TabIndex        =   2
      Top             =   9120
      Width           =   2055
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
      Left            =   16680
      TabIndex        =   1
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label lblJason 
      Alignment       =   2  'Center
      Caption         =   "Jason"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8280
      TabIndex        =   0
      Top             =   11640
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   9930
      Left            =   3120
      Picture         =   "PicturesJason.frx":0000
      Top             =   1320
      Width           =   13245
   End
End
Attribute VB_Name = "frmPicturesJason"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Address Book
'frmPicturesJason(PicturesJason.frm)
'Sarah Keating
'10-27-03
'Purpose: This form shows a picture of Jason.

Private Sub cmdNext_Click()
frmPicturesJohn.Show
frmPicturesJason.Hide
' Allows the User to go to the next photo
End Sub

Private Sub cmdPrevious_Click()
frmPicturesJason.Hide
frmPicturesPat.Show
' Allows the user to go to the previous photo
End Sub

Private Sub cmdQuit_Click()
    End
    ' Allows the user to quit the program
End Sub
