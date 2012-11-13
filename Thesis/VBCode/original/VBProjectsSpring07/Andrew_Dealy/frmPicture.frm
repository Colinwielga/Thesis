VERSION 5.00
Begin VB.Form frmPicture 
   BackColor       =   &H00000000&
   Caption         =   "Andrew Dealy"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to main page"
      Height          =   495
      Left            =   6600
      TabIndex        =   2
      Top             =   5880
      Width           =   2055
   End
   Begin VB.PictureBox picPicture 
      BackColor       =   &H00000000&
      Height          =   6015
      Left            =   240
      Picture         =   "frmPicture.frx":0000
      ScaleHeight     =   5955
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
   Begin VB.Label lblBiography 
      BackColor       =   &H80000007&
      Caption         =   $"frmPicture.frx":7AD8
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6015
      Left            =   5520
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "frmPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this page is a short bio on the creator with a link to the main page
Private Sub cmdReturn_Click()
    frmShaunWhite.Show
    frmPicture.Hide
End Sub
