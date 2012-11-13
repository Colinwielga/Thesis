VERSION 5.00
Begin VB.Form frmPicturesKatieF 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Katie"
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
      Left            =   15960
      TabIndex        =   3
      Top             =   12600
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
      Height          =   2055
      Left            =   15960
      TabIndex        =   2
      Top             =   10080
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
      Left            =   15960
      TabIndex        =   1
      Top             =   7560
      Width           =   1935
   End
   Begin VB.Label lblKatieF 
      Alignment       =   2  'Center
      Caption         =   "Katie Furniss"
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
      Left            =   9120
      TabIndex        =   0
      Top             =   12480
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   11685
      Left            =   6120
      Picture         =   "PicturesKatieF.frx":0000
      Top             =   480
      Width           =   8775
   End
End
Attribute VB_Name = "frmPicturesKatieF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Address Book
'frmPicturesKatieF(PicturesKatieF.frm)
'Sarah Keating
'10-27-03
'Purpose: This form shows a picture of Katie Furniss.

Private Sub cmdNext_Click()
frmPicturesPat.Show
frmPicturesKatieF.Hide
' Allows the User to go to the next photo
End Sub

Private Sub cmdPrevious_Click()
frmPicturesKatieF.Hide
frmPicturesJessi.Show
' Allows the user to go to the previous photo
End Sub

Private Sub cmdQuit_Click()
    End
    ' Allows the user to quit the program
End Sub
